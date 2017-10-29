#!/usr/bin/env python

import argparse
import datetime
import os
import json
import plistlib
import re
import sqlite3
import subprocess
import sys
import time
import urllib2


def z_get_userid(token):
    """Validate the Zotero API token and return the user ID"""
    req = urllib2.Request("https://api.zotero.org/keys/%s" % token)
    req.add_header("Zotero-API-Version", "3")
    try:
        f = urllib2.urlopen(req)
    except urllib2.HTTPError:
        sys.exit("Could not authenticate with Zotero API: was the token "
                 "correct?")
    result = json.load(f)
    if not (result["access"]["user"]["library"]
            and result["access"]["user"]["write"]):
        sys.exit("This token does not appear to have sufficient privileges")
    return result["userID"]


def open_papersdb():
    """Return the connection cursor to the papers sqlite library"""
    f = os.path.expanduser("~/Library/Preferences/com.mekentosj.papers3.plist")

    # Unfortunately, plistlib does not handle binary plists so the plutil
    # shell utility is needed to avoid depending on external libraries
    plist = subprocess.check_output(["plutil", "-convert", "xml1", "-o", "-", f])
    lib = plistlib.readPlistFromString(plist)["mt_papers3_library_location_local"]
    sqlfile = "/".join([os.path.expanduser("~/Library/Application Support"),
                        lib,
                        "Library.papers3/Database.papersdb"
                        ])

    conn = sqlite3.connect(sqlfile)
    conn.row_factory = sqlite3.Row
    return conn.cursor()


def z_recreate_collections(token, userid, papersdb_cursor):
    """Recreate the collection structure from papers

    This function returns a dictionary mapping the papers collection UUID
    to the Zotero API key.
    """

    # List all top-level collections and generate a unique name for the
    # library import
    tlds_url = "https://api.zotero.org/users/%s/collections/top" % userid
    tlds_req = urllib2.Request(tlds_url)
    tlds_req.add_header("Zotero-API-Key", token)
    tlds_req.add_header("Zotero-API-Version", "3")
    tlds = [x["data"]["name"] for x in json.load(urllib2.urlopen(tlds_req))]

    now = datetime.datetime.utcnow().replace(microsecond=0).isoformat()
    new_tld = "_".join(["passport-import", now])
    while new_tld in tlds:
        time.sleep(1)
        now = datetime.datetime.utcnow().replace(microsecond=0).isoformat()
        new_tld = "_".join(["passport-import", now])

    # Create a passport import collection and record the zotero ID
    lib_data = json.dumps([{"name": new_tld,
                            "parentCollection": False}])
    lib_url = "https://api.zotero.org/users/%s/collections" % userid
    lib_req = urllib2.Request(lib_url, lib_data)
    lib_req.add_header("Zotero-API-key", token)
    lib_req.add_header("Zotero-API-Version", "3")
    lib_req.add_header("Content-Type", "application/json")
    lib_res = json.load(urllib2.urlopen(lib_req))
    if not "success" in lib_res:
        sys.exit("Could not create a new collection for import")
    p_tld_sql = ("SELECT uuid FROM Collection WHERE editable = 0 "
                "AND name = 'COLLECTIONS';")
    papersdb_cursor.execute(p_tld_sql)
    p_tld_uuid = papersdb_cursor.fetchone()[0]
    collection_map = {p_tld_uuid: lib_res["success"]["0"]}
    # For ease of later referencing in z_recreate_items for orphaned items, we
    # also insert it into the map with key "tld"
    collection_map["tld"] = lib_res["success"]["0"]
    
    # Get a list of all collections in papers
    p_sql = ("SELECT uuid, name, parent FROM Collection WHERE editable=1")
    p_collections = {}
    for row in papersdb_cursor.execute(p_sql):
        p_collections[row[0]] = {"name": row[1], "parent": row[2]}

    # Create all level 1 collections (i.e. not sub-collections)
    for uuid, v in p_collections.copy().iteritems():
        if v["parent"] == p_tld_uuid:
            level1_data = json.dumps([{
                "name": v["name"],
                "parentCollection": collection_map[p_tld_uuid]
                }])
            level1_url = "https://api.zotero.org/users/%s/collections" \
                         % userid
            level1_req = urllib2.Request(level1_url, level1_data)
            level1_req.add_header("Zotero-API-Key", token)
            level1_req.add_header("Zotero-API-Version", "3")
            level1_req.add_header("Content-Type", "application/json")
            level1_res = urllib2.urlopen(level1_req)
            collection_map[uuid] = json.load(level1_res)["success"]["0"]
            del p_collections[uuid]

    # Create all level >1 collections by looping through 
    # collection_map and searching p_collections for children
    while len(p_collections) > 0:
        for p_uuid in collection_map.copy().iterkeys():
            to_add = {k: v for k, v in p_collections.items()
                                    if v["parent"] == p_uuid}
            for uuid, v in to_add.iteritems():
                add_data = json.dumps([{
                    "name": v["name"],
                    "parentCollection": collection_map[v["parent"]]
                    }])
                add_url = "https://api.zotero.org/users/%s/collections" \
                          % userid
                add_req = urllib2.Request(add_url, add_data)
                add_req.add_header("Zotero-API-Key", token)
                add_req.add_header("Zotero-API-Version", "3")
                add_req.add_header("Content-Type", "application/json")
                add_res = urllib2.urlopen(add_req)
                collection_map[uuid] = json.load(add_res)["success"]["0"]
                del p_collections[uuid]

    return collection_map


def z_recreate_items(token, userid, papersdb_cursor, collection_map):
    """Import items from papers into the Zotero API"""
    items_sql = ("SELECT "
                 "a.uuid AS uuid, "
                 "a.title AS title, "
                 "b.abbreviation AS journalAbbreviation, "
                 "b.title AS journalTitle, "
                 "a.volume AS volume, "
                 "a.number AS number, "
                 "a.startpage AS startpage, "
                 "a.endpage AS endpage, "
                 "a.publication_date AS publication_date, "
                 "a.language AS language, "
                 "a.doi AS doi, "
                 "a.imported_date AS imported_date, "
                 "a.notes AS notes "
                 "FROM Publication a, Publication b "
                 "WHERE a.bundle = b.uuid "
                 "AND a.type >= 0 "
                 "AND a.privacy_level = 0 LIMIT 500")
    items_res = papersdb_cursor.execute(items_sql)
    import_items = []
    import_notes = []
    import_pubmed = []
    for item in items_res.fetchall():
        jsondict = {"itemType": "journalArticle"}

        if item["title"] is not None:
            jsondict["title"] = item["title"]
        
        # Parse authors into list of dictionaries as per Zotero API
        authors_sql = ("SELECT Author.prename, Author.surname "
                       "FROM OrderedAuthor "
                       "LEFT JOIN Author "
                       "ON OrderedAuthor.author_id = Author.uuid "
                       "WHERE OrderedAuthor.object_id = ? "
                       "AND OrderedAuthor.type = 0 "
                       "ORDER BY OrderedAuthor.priority;")
        authors_res = papersdb_cursor.execute(authors_sql, (item["uuid"],))
        authors = []
        for author in authors_res.fetchall():
            firstName = author["prename"]
            # Zotero puts a period after each initial
            while (re.search(r"(^| )[A-Z]( |$)", firstName) is not None):
                firstName = re.sub(r"((^| )[A-Z])( |$)", r"\1.\3",
                    firstName)
            authors.append({
                "creatorType": "author",
                "firstName": firstName,
                "lastName": author["surname"]
                })
        if len(authors) > 0:
            jsondict["creators"] = authors

        if item["journalTitle"] is not None:
            jsondict["publicationTitle"] = item["journalTitle"]

        if item["journalAbbreviation"] is not None:
            jsondict["journalAbbreviation"] = item["journalAbbreviation"]

        if item["volume"] is not None:
            jsondict["volume"] = item["volume"]

        if item["number"] is not None:
            jsondict["issue"] = item["number"]

        if item["startpage"] is not None:
            if item["endpage"] is None:
                jsondict["pages"] = item["startpage"]
            else:
                jsondict["pages"] = "-".join([
                    item["startpage"],
                    item["endpage"]
                    ])

        # Format the publication date to the correct degree of accuracy
        if len(item["publication_date"][2:10]) == 8:
            if int(item["publication_date"][2:6]) > 0:
                publication_date = item["publication_date"][2:6]
                if int(item["publication_date"][6:8]) > 0:
                    publication_date = "-".join([
                        publication_date,
                        item["publication_date"][6:8]
                        ])
                    if int(item["publication_date"][8:10]) > 0:
                        publication_date = "-".join([
                            publication_date,
                            item["publication_date"][8:10]
                            ])
        jsondict["date"] = publication_date

        if item["language"] is not None:
            jsondict["language"] = item["language"]

        if item["doi"] is not None:
            jsondict["doi"] = item["doi"]

        dateAdded = datetime.datetime.utcfromtimestamp(item["imported_date"])
        dateAdded = "".join([dateAdded.replace(microsecond=0).isoformat(), "Z"])
        jsondict["dateAdded"] = dateAdded

        # PMID / PMC
        pubmed_sql = ("SELECT remote_id, source_id from SyncEvent "
                      "WHERE device_id = ? "
                      "AND subtype = 0 "
                      "AND (source_id = 'gov.nih.nlm.ncbi.pubmed' "
                      "OR source_id = 'gov.nih.nlm.ncbi.pmc');")
        pubmed_res = papersdb_cursor.execute(pubmed_sql, (item["uuid"],))
        extra = []
        for pubmed_row in pubmed_res.fetchall():
            if pubmed_row["source_id"] == "gov.nih.nlm.ncbi.pubmed":
                extra.append("PMID: %s" % pubmed_row["remote_id"])

                # PubMed entry attachment
                import_pubmed.append({
                    "itemType": "attachment",
                    "linkMode": "linked_url",
                    "title": "PubMed entry",
                    "accessDate": dateAdded,
                    "url": "http://www.ncbi.nlm.nih.gov/pubmed/%s" %
                        pubmed_row["remote_id"],
                    "note": "",
                    "contentType": "text/html",
                    "tags": [],
                    "collections": [],
                    "relations": {},
                    "charset": "",
                    "papers_uuid": item["uuid"]
                    })
            elif pubmed_row["source_id"] == "gov.nih.nlm.ncbi.pmc":
                extra.append("PMCID: %s" % pubmed_row["remote_id"])
        if len(extra) > 0:
            jsondict["libraryCatalog"] = "PubMed"
            jsondict["extra"] = "\n".join(extra)

        # Tags
        tags_sql = ("SELECT Keyword.name FROM KeywordItem "
                    "LEFT JOIN Keyword "
                    "ON KeywordItem.keyword_id = Keyword.uuid "
                    "WHERE KeywordItem.object_id = ? "
                    "AND KeywordItem.type = 99;")
        tags_res = papersdb_cursor.execute(tags_sql, (item["uuid"],))
        tags = []
        for tag in tags_res.fetchall():
            tags.append({"tag": tag["name"], "type": 1})
        jsondict["tags"] = tags

        # Collections need to be mapped from their papers uuid to the zotero key
        coll_sql = ("SELECT Collection.uuid FROM CollectionItem "
                    "LEFT JOIN Collection "
                    "ON CollectionItem.collection = Collection.uuid "
                    "WHERE CollectionItem.object_id = ?;")
        coll_res = papersdb_cursor.execute(coll_sql, (item["uuid"],))
        collections = []
        for coll in coll_res.fetchall():
            # For some reason, some items in papers are assigned to a collection
            # that does not exist: put these in the top level import folder
            if coll["uuid"] is not None:
                collections.append(collection_map[coll["uuid"]])
            else:
                collections.append(collection_map["tld"])
        jsondict["collections"] = collections

        jsondict["relations"] = {}

        # Notes can only be imported once we have the key to the parent item.
        # They are therefore added to a separate list `import_notes` which also
        # contains the dictionary item "papers_uuid".  Later, when we have the
        # item_map we replace papers_uuid with parentItem and upload them.
        if item["notes"] is not None:
            import_notes.append({
                "itemType": "note",
                "note": item["notes"],
                "tags": [],
                "collections": [],
                "relations": {},
                "papers_uuid": item["uuid"]
                })

        # Although the papers uuid does not need importing into zotero,
        # add it to the array so that the item_map can be built later
        jsondict["papers_uuid"] = item["uuid"]
        
        # Add this item to the `import_items` list for importing
        import_items.append(jsondict)

    # Upload items in batches of 50
    import_items_chunks = [import_items[i:i+50]
                           for i in xrange(0, len(import_items), 50)]
    item_map = {}
    for chunk in import_items_chunks:
        item_map_chunk = {}
        i = 0
        for item in chunk:
            item_map_chunk[i] = item["papers_uuid"]
            del chunk[i]["papers_uuid"]
            i += 1

        items_url = "https://api.zotero.org/users/%s/items" % userid
        items_req = urllib2.Request(items_url, json.dumps(chunk))
        items_req.add_header("Zotero-API-Key", token)
        items_req.add_header("Zotero-API-Version", "3")
        items_req.add_header("Content-Type", "application/json")
        items_res = urllib2.urlopen(items_req)

        items_success = json.load(items_res)["success"]
        for x in items_success:
            item_map[item_map_chunk[int(x)]] = items_success[x]

    # Upload notes
    for note_key, note in enumerate(import_notes):
        import_notes[note_key]["parentItem"] = item_map[note["papers_uuid"]]
        del import_notes[note_key]["papers_uuid"]
    import_notes_chunks = [import_notes[i:i+50]
                           for i in xrange(0, len(import_notes), 50)]
    for chunk in import_notes_chunks:
        notes_url = "https://api.zotero.org/users/%s/items" % userid
        notes_req = urllib2.Request(notes_url, json.dumps(chunk))
        notes_req.add_header("Zotero-API-Key", token)
        notes_req.add_header("Zotero-API-Version", "3")
        notes_req.add_header("Content-Type", "application/json")
        urllib2.urlopen(notes_req)

    # Upload PubMed entries
    for pubmed_key, pubmed in enumerate(import_pubmed):
        import_pubmed[pubmed_key]["parentItem"] = item_map[pubmed["papers_uuid"]]
        del import_pubmed[pubmed_key]["papers_uuid"]
    import_pubmed_chunks = [import_pubmed[i:i+50]
                            for i in xrange(0, len(import_pubmed), 50)]
    for chunk in import_pubmed_chunks:
        pubmed_url = "https://api.zotero.org/users/%s/items" % userid
        pubmed_req = urllib2.Request(pubmed_url, json.dumps(chunk))
        pubmed_req.add_header("Zotero-API-Key", token)
        pubmed_req.add_header("Zotero-API-Version", "3")
        pubmed_req.add_header("Content-Type", "application/json")
        urllib2.urlopen(pubmed_req)

    return item_map


def main():
    description = """
    Import a Papers 3 library to Zotero.  For more information see:
    https://andrewlkho.github.com/passport.
    """
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument("--token",
                        help="Specify API key")
    args = parser.parse_args()

    papersdb_cursor = open_papersdb()
    userid = z_get_userid(args.token)
    collection_map = z_recreate_collections(args.token, userid, papersdb_cursor)
    z_recreate_items(args.token, userid, papersdb_cursor, collection_map)


if __name__ == "__main__":
    main()
