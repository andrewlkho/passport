Passport is a utility for transferring a library from
[Papers 3](https://www.readcube.com/papers/) to
[Zotero](https://www.zotero.org/).

There already exist several ways of transferring a list of papers such as
exporting to and then importing from a RIS file.  However, none of them preserve
the collections I have sorted my papers into.

Passport will:
- Recreate your Papers 3 collections and file your papers as appropriate
- Copy across any notes and tags from Papers 3
- Copy the PDFs to your local Zotero data storage directory
- Optionally, pass your library through PubMed to clean up some of the metadata
  (see below)
- Not make any changes to your Papers 3 library which it only reads data from

As the Papers 3 sqlite library is not a documented format, I have had to reverse
engineer the library format.  As such, I can make no guarantees that passport
will correctly read and transfer across your information; all I can say is that
it works on my library.  I would strongly recommend that you at least check that
the correct number of papers have been copied across to each collection as well
as keep a backup of your Papers 3 library for future reference until you are
satisfied.


# Usage

1. Log in to Zotero and create a new API key for passport
   ([Settings > Feeds/API > New Key](https://www.zotero.org/settings/keys/new)).
   The new key will need to have full access to your personal library.  Make
   a note of the key produced.

2. Open Terminal.app

3. Download the passport script and make it executable:

    ```
    % curl -O
    https://raw.githubusercontent.com/andrewlkho/passport/master/passport.py
    % chmod +x ./passport.py
    ```

4. Run the script, passing it your API key (`<KEY>` in the command below):

    % ./passport.py --token <KEY>


# Cleaning metadata through PubMed

Passport can optionally replace the journal title and/or abstract with data from
PubMed where a match can be found.  You can do this by passing `--pubmed-cleanup
journal` and/or `--pubmed-cleanup abstract` when invoking the script.  Note that
doing so will also update the PMID, PMCID and DOI if available.  Note also that
this is a direct overwrite, so if for whatever reason you have manually edited
this data your changes will be lost.

Why the journal title?  For some reason I have very inconsistent naming in my
Papers 3 library.  For example, I have the [Red 
Journal](http://www.redjournal.org/) as

- IJROBP
- Int. J. Radiat. Oncol. Biol. Phys.
- International Journal of Radiation Oncology\*Biology\*Physics
- International Journal of Radiation Oncology, Biology, Physics


# Doesn't work for you?

If passport doesn't transfer your library then do
[file an issue](https://github.com/andrewlkho/passport/issues).
