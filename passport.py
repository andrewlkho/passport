#!/usr/bin/env python

import argparse
import ConfigParser
import datetime
import os
import json
import plistlib
import re
import shutil
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


def z_api_write(token, url, data):
    """Write the supplied data to the zotero API.

    The data should be in the format of a list of dictionaries which is
    converted into the JSON expected by the Zotero API.  It takes care of
    uploading in batches of fifty.  It returns the "success" dictionary where
    the key is the list index corresponding to that item, and the value is
    the zotero key returned.
    """

    data_chunks = [data[i:i+50] for i in xrange(0, len(data), 50)]
    success = {}
    for chunk_key, chunk in enumerate(data_chunks):
        req = urllib2.Request(url, json.dumps(chunk))
        req.add_header("Zotero-API-Key", token)
        req.add_header("Zotero-API-Version", "3")
        req.add_header("Content-Type", "application/json")
        try:
            res = urllib2.urlopen(req)
        except urllib2.HTTPError as e:
            sys.exit("Error: received HTTP %s" % e.code)
        else:
            # if len(json.load(res)["failed"]) > 0:
            #     sys.exit("Error: failed to write")
            res_success = json.load(res)["success"]
            for k in res_success.iterkeys():
                success[(chunk_key * 50) + int(k)] = res_success[k]
    return success


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
    lib_data = [{"name": new_tld,
                "parentCollection": False}]
    lib_res = z_api_write(
            token,
            "https://api.zotero.org/users/%s/collections" % userid,
            lib_data
            )
    try:
        tld_key = lib_res[0]
    except KeyError:
        sys.exit("Could not create a new collection for import")
    p_tld_sql = ("SELECT uuid FROM Collection WHERE editable = 0 "
                "AND name = 'COLLECTIONS';")
    papersdb_cursor.execute(p_tld_sql)
    p_tld_uuid = papersdb_cursor.fetchone()[0]
    collection_map = {p_tld_uuid: tld_key}
    # For ease of later referencing in z_recreate_items for orphaned items, we
    # also insert it into the map with key "tld"
    collection_map["tld"] = tld_key

    # Get a list of all collections in papers
    p_sql = ("SELECT uuid, name, parent FROM Collection WHERE editable=1")
    p_collections = {}
    for row in papersdb_cursor.execute(p_sql):
        p_collections[row[0]] = {"name": row[1], "parent": row[2]}

    # Create all level 1 collections (i.e. not sub-collections)
    for uuid, v in p_collections.copy().iteritems():
        if v["parent"] == p_tld_uuid:
            level1_data = [{
                "name": v["name"],
                "parentCollection": collection_map[p_tld_uuid]
                }]
            level1_res = z_api_write(
                    token,
                    "https://api.zotero.org/users/%s/collections" % userid,
                    level1_data
                    )
            collection_map[uuid] = level1_res[0]
            del p_collections[uuid]

    # Create all level >1 collections by looping through 
    # collection_map and searching p_collections for children
    while len(p_collections) > 0:
        for p_uuid in collection_map.copy().iterkeys():
            to_add = {k: v for k, v in p_collections.items()
                                    if v["parent"] == p_uuid}
            for uuid, v in to_add.iteritems():
                add_data = [{
                    "name": v["name"],
                    "parentCollection": collection_map[v["parent"]]
                    }]
                add_res = z_api_write(
                        token,
                        "https://api.zotero.org/users/%s/collections" % userid,
                        add_data
                        )
                collection_map[uuid] = add_res[0]
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

    # Upload items
    item_map = {}
    item_map_uuids = {}
    for i, item in enumerate(import_items):
        item_map_uuids[i] = item["papers_uuid"]
        del import_items[i]["papers_uuid"]
    item_success = z_api_write(
            token,
            "https://api.zotero.org/users/%s/items" % userid,
            import_items
            )
    for x in item_success:
        item_map[item_map_uuids[int(x)]] = item_success[x]

    # Upload notes
    for i, note in enumerate(import_notes):
        import_notes[i]["parentItem"] = item_map[note["papers_uuid"]]
        del import_notes[i]["papers_uuid"]
    z_api_write(
            token,
            "https://api.zotero.org/users/%s/items" % userid,
            import_notes
            )

    # Upload PubMed entries
    for i, pubmed in enumerate(import_pubmed):
        import_pubmed[i]["parentItem"] = item_map[pubmed["papers_uuid"]]
        del import_pubmed[i]["papers_uuid"]
    z_api_write(
            token,
            "https://api.zotero.org/users/%s/items" % userid,
            import_pubmed
            )

    return item_map


def z_recreate_pdfs(token, userid, papersdb_cursor, item_map):
    """Copy PDFs to zotero local storage and upload info to API"""
    # Get path to zotero data directory
    config = ConfigParser.RawConfigParser()
    profilesini = config.read(os.path.expanduser(
        "~/Library/Application Support/Zotero/profiles.ini"
        ))
    prefsjs = open(os.path.expanduser("/".join([
        "~/Library/Application Support/Zotero",
        config.get("Profile0", "Path"),
        "prefs.js"
        ])))
    for line in prefsjs:
        match = re.search(
                r'user_pref\("extensions\.zotero\.dataDir", "([^"]+)"\);',
                line
                )
        if match:
            datadir = match.group(1)

    # Generate "pdfs", a list where each item is a dictionary containing
    # a PDF path, a zotero key and the date the item was added
    f = os.path.expanduser("~/Library/Preferences/com.mekentosj.papers3.plist")
    plist = subprocess.check_output(["plutil", "-convert", "xml1", "-o", "-", f])
    prefix = plistlib.readPlistFromString(plist)[
                 "mt_papers3_full_library_location_shared"
                 ]
    pdfs_sql = ("SELECT path, object_id, created_at FROM PDF "
                "WHERE type = 0 "
                "AND mime_type = 'application/pdf';")
    pdfs_res = papersdb_cursor.execute(pdfs_sql)
    pdfs = []
    for pdfs_row in pdfs_res.fetchall():
        path_abs = "/".join([prefix, pdfs_row["path"]])
        if pdfs_row["object_id"] in item_map:
            if os.path.isfile(path_abs):
                d = datetime.datetime.utcfromtimestamp(pdfs_row["created_at"])
                d = "".join([d.replace(microsecond=0).isoformat(), "Z"])
                pdfs.append({
                        "path": path_abs,
                        "parentItem": item_map[pdfs_row["object_id"]],
                        "dateAdded": d
                        })

    # Create import_pdfs and upload to the zotero API
    import_pdfs = []
    for pdf in pdfs:
        import_pdfs.append({
            "itemType": "attachment",
            "linkMode": "imported_file",
            "title": os.path.basename(pdf["path"]),
            "contentType": "application/pdf",
            "filename": os.path.basename(pdf["path"]),
            "tags": [],
            "relations": {},
            "dateAdded": pdf["dateAdded"],
            "parentItem": pdf["parentItem"]
            })
    pdfs_success = z_api_write(
            token,
            "https://api.zotero.org/users/%s/items" % userid,
            import_pdfs
            )

    # Copy PDFs to the zotero data directory
    for i in pdfs_success:
        dest = "/".join([datadir, "storage", pdfs_success[i]])
        os.mkdir(dest)
        shutil.copy(pdfs[i]["path"], dest)


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
    item_map = z_recreate_items(args.token, userid, papersdb_cursor,
            collection_map)
    z_recreate_pdfs(args.token, userid, papersdb_cursor, item_map)


if __name__ == "__main__":
    main()
