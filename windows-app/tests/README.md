# Storage smoke check

On a development machine with Mono, compile the store and its smoke test:

```sh
mcs -sdk:2 -out:/tmp/PJShoesStoreSmokeTest.exe \
  -r:System -r:System.Core -r:System.Xml -r:System.Xml.Linq -r:WindowsBase \
  src/SlipRecord.cs src/SpreadsheetStore.cs src/LegacyWorkbookImporter.cs \
  tests/StoreSmokeTest.cs
mono /tmp/PJShoesStoreSmokeTest.exe
```

Pass a small `.xlsx` fixture as the first argument to also exercise the legacy
workbook importer. This test does not replace the required Windows 7 RTM and
printer smoke test.
