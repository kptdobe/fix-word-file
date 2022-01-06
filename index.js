import fetch from 'node-fetch';
import fs from 'fs-extra';
import dotenv from 'dotenv';

dotenv.config();

async function repairWordDoc(site, filePath, cookie) {
  let xml = await fs.readFile('repair.xml');
  xml = xml.toString().replace('${url}', `${site}${filePath}`);

  let body = '\r\n';
  body += '--urn:uuid:D5FEDF9B-4A2A-4841-83BD-5E670DE4B6A8\r\n';
  body += 'Content-ID: <http://tempuri.org/xml/0>\r\n';
  body += 'Content-Transfer-Encoding: 8bit\r\n';
  body += 'Content-Type: application/xop+xml;charset=utf-8;type="text/xml; charset=utf-8"\r\n';
  body += '\r\n';
  body += `${xml}\r\n`;
  body += '--urn:uuid:D5FEDF9B-4A2A-4841-83BD-5E670DE4B6A8--';

  const api = `${site}/_vti_bin/cellstorage.svc/CellStorageService`;
  const headers = {
    'Cookie': cookie,
    'SOAPAction': '"http://schemas.microsoft.com/sharepoint/soap/ICellStorages/ExecuteCellStorageRequest"',
    'Bearer': 'Authorization',
    'Accept-Auth': 'badger,Wlid1.1,Bearer,Basic,NTLM,Digest,Kerberos,Negotiate,Nego2',
    'Content-Type': 'multipart/related; type="application/xop+xml"; boundary="urn:uuid:D5FEDF9B-4A2A-4841-83BD-5E670DE4B6A8"; start="<http://tempuri.org/xml/0>"; start-Info="text/xml; charset=utf-8"',
    'MIME-Version': '1.0',
  };

  const response = await fetch(api, {
    method: 'POST',
    body,
    headers,
  });

  if (response.status === 200) {
    console.log(`Document ${filePath} should be repaired.`);
  } else {
    const msg = response.text();
    console.error(`Document ${filePath} could NOT be repaired: ${msg}`);
  }
}
async function main() {
  const args = process.argv.slice(2);
  if (args.length !== 2) {
    console.error('Usage: node index.js <url to the site> <full path to the file in the site>');
  } else {
    const site = args[0];
    const filePath = args[1];
    await repairWordDoc(site, filePath, process.env.WORD_COOKIE);
  }
  
}

main();

