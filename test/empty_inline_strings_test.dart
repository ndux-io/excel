import 'dart:convert';
import 'dart:io';

import 'package:archive/archive.dart';
import 'package:excel/excel.dart';
import 'package:test/test.dart';
import 'package:xml/xml.dart';

void main() {
  group('empty inline strings', () {
    test('decodes empty cells and surrounding values from bytes', () {
      _expectCells(Excel.decodeBytes(_workbookBytes()));
    });

    test('decodes empty cells from a memory buffer', () {
      _expectCells(Excel.decodeBuffer(InputMemoryStream(_workbookBytes())));
    });

    test('decodes empty cells from a file buffer', () {
      final directory = Directory.systemTemp.createTempSync('empty_inline_');
      try {
        final file = File('${directory.path}/empty.xlsx')
          ..writeAsBytesSync(_workbookBytes());
        final input = InputFileStream(file.path);
        try {
          _expectCells(Excel.decodeBuffer(input));
        } finally {
          input.closeSync();
        }
      } finally {
        directory.deleteSync(recursive: true);
      }
    });

    test('preserves blank cell styles and values after round-trip', () {
      final excel = Excel.decodeBytes(_workbookBytes());
      _expectCells(Excel.decodeBytes(excel.encode()!));
    });

    test('decodes empty cells with namespace-prefixed SpreadsheetML', () {
      final archive = ZipDecoder().decodeBytes(_workbookBytes());
      for (final file
          in archive.files.where((file) => file.name.endsWith('.xml'))) {
        final document = XmlDocument.parse(utf8.decode(file.content));
        if (document.rootElement.namespaceUri != _spreadsheetNamespace)
          continue;
        final root = _prefixElements(document.rootElement) as XmlElement;
        root.setAttribute('xmlns:x', _spreadsheetNamespace);
        archive.addFile(ArchiveFile.string(file.name, root.toXmlString()));
      }
      final bytes = ZipEncoder().encode(archive);
      _expectCells(Excel.decodeBytes(bytes));
      _expectCells(Excel.decodeBuffer(InputMemoryStream(bytes)));
    });

    test('does not suppress missing values for shared strings', () {
      expect(
        () => Excel.decodeBytes(_workbookBytes(extraCell: '<c r="I2" t="s"/>')),
        throwsStateError,
      );
    });
  });
}

const _spreadsheetNamespace =
    'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

XmlNode _prefixElements(XmlNode node) {
  if (node is! XmlElement) return node.copy();
  return XmlElement(
    node.namespaceUri == _spreadsheetNamespace
        ? XmlName.parts(node.name.local, prefix: 'x')
        : node.name,
    node.attributes.map((attribute) => attribute.copy()),
    node.children.map(_prefixElements),
    node.isSelfClosing,
  );
}

void _expectCells(Excel excel) {
  final sheet = excel.tables['Sheet1']!;
  expect(sheet.maxRows, 2);
  expect(sheet.maxColumns, 8);
  expect(sheet.rows[0][0]?.value, TextCellValue('Header'));
  expect(sheet.rows[1][0]?.value, TextCellValue('질문 & 내용'));
  for (final column in [1, 2, 3]) {
    expect(sheet.rows[1][column]?.value, isNull);
    expect(sheet.rows[1][column]?.cellStyle?.isBold, isTrue);
  }
  expect(sheet.rows[1][4]?.value, TextCellValue(''));
  expect(sheet.rows[1][5]?.value, TextCellValue('Text'));
  expect(sheet.rows[1][6]?.value, IntCellValue(42));
  expect(sheet.rows[1][7]?.value, TextCellValue('After'));
}

List<int> _workbookBytes({String extraCell = ''}) {
  final excel = Excel.createExcel();
  excel['Sheet1'].updateCell(
    CellIndex.indexByString('A1'),
    TextCellValue('Header'),
    cellStyle: CellStyle(bold: true),
  );
  final archive = ZipDecoder().decodeBytes(excel.encode()!);
  const path = 'xl/worksheets/sheet1.xml';
  final original =
      XmlDocument.parse(utf8.decode(archive.findFile(path)!.content));
  final style = original.findAllElements('c').first.getAttribute('s')!;
  archive.addFile(ArchiveFile.string(path, '''
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Header</t></is></c></row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>질문 &amp; 내용</t></is></c>
      <c r="B2" s="$style" t="inlineStr"></c>
      <c r="C2" s="$style" t="inlineStr"/>
      <c r="D2" s="$style" t="inlineStr"><is/></c>
      <c r="E2" t="inlineStr"><is><t/></is></c>
      <c r="F2" t="inlineStr"><is><t>Text</t></is></c>
      <c r="G2" t="n"><v>42</v></c>
      <c r="H2" t="inlineStr"><is><t>After</t></is></c>
      $extraCell
    </row>
  </sheetData>
</worksheet>
'''));
  return ZipEncoder().encode(archive);
}
