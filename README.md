# fix-word-file

Help fixing a Word file in Sharepoint which cannot be opened via the Web UI but requires to be opened with the Word desktop app and no modification is made to the file.
## Cookie

To connect to the server, you need to provide a cookie. To get the cookie, you need to snif the Word requests and extract it from there.
Follow the instructions from https://earthly.dev/blog/mitmproxy/ to track all request your (Mac) computer is making.
Then, when opening a Word file from Sharepoint with the Word desktop app, just filter the request to your Sharepoint instance, some of them will have a cookie. Copy the value and put it in a `.env` file:

```
WORD_COOKIE=SPOIDCRL=77u/PD94bWwgdm...............vU1A+
```

## Usage

```
node index.js <url to the site> <full path to the file in the site>
```

Example:

```
node index.js https://adobe.sharepoint.com/sites/AlexSite /Shared%20Documents/a/b/c/myfile.docx;
```

## Notes

The repair.xml might probably contain some identifier specific to the Sharepoint instance / site I am using. You can copy one from the request Word is making if this one does not work. 

The following subrequest seems to be the "thing" that fixes the file. No clue what it does.

```
<SubRequest
    Type="Cell"
    SubRequestToken="11">
  <SubRequestData BinaryDataSize="27">DgALAJzPKfM5lAabBgIAABYCBgADAwALAQMB</SubRequestData>
</SubRequest>
```