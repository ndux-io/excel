import 'dart:convert';
import 'package:archive/archive.dart';
import 'package:excel/excel.dart';
import 'package:test/test.dart';
import 'package:xml/xml.dart';

void main() {
  for (final buffer in [false, true]) {
    Excel decode(List<int> bytes) => buffer
        ? Excel.decodeBuffer(InputMemoryStream(bytes))
        : Excel.decodeBytes(bytes);
    group(buffer ? 'buffer' : 'bytes', () {
      test('inline and shared runs preserve ordering, whitespace and emphasis',
          () {
        const runs = '<t>앞 </t><r><rPr><b/></rPr><t>중간</t></r>'
            '<t> 뒤</t><r><rPr><b val="0"/></rPr><t>끝</t></r>';
        final workbook = decode(fixture(
            '<c r="A1" t="inlineStr"><is>$runs</is></c>'
            '<c r="B1" t="s"><v>0</v></c>',
            shared: runs));
        for (final cell in workbook.tables.values.first.rows.first.take(2)) {
          expect(cell!.value.toString(), '앞 중간 뒤끝');
          final span = (cell.value as TextCellValue).value;
          expect(span.children![1].style!.isBold, isTrue);
          expect(span.children![3].style!.isBold, isFalse);
        }
      });
      test('ISO date, date-time and fraction preserve wall-clock fields', () {
        final workbook = decode(fixture('<c r="A1" t="d"><v>2026-09-06</v></c>'
            '<c r="B1" t="d"><v>2026-09-06T12:30:01.123456+09:00</v></c>'));
        final cells = workbook.tables.values.first.rows.first;
        expect(
            cells[0]!.value, const DateCellValue(year: 2026, month: 9, day: 6));
        expect(
            cells[1]!.value,
            const DateTimeCellValue(
                year: 2026,
                month: 9,
                day: 6,
                hour: 12,
                minute: 30,
                second: 1,
                millisecond: 123,
                microsecond: 456));
      });
      test('invalid ISO dates reject the workbook', () {
        for (final value in ['bad', '2026-02-30', '2026-09-06T25:00:00']) {
          expect(() => decode(fixture('<c r="A1" t="d"><v>$value</v></c>')),
              throwsFormatException);
        }
      });
      test('workbook date system is exposed and defaults to false', () {
        expect(decode(fixture('<c r="A1"/>')).uses1904DateSystem, isFalse);
        for (final value in ['1', 'true']) {
          expect(
              decode(fixture('<c r="A1"/>', date1904: value))
                  .uses1904DateSystem,
              isTrue);
        }
        expect(decode(fixture('<c r="A1"/>', date1904: '0')).uses1904DateSystem,
            isFalse);
      });
      test('foreign numeric and date formats decode to the correct types', () {
        for (final format in [
          '[Red]-#,##0.00',
          '[>=100]0.00',
          '0 "days"',
          r'0\d',
          '0_d',
          '0*d'
        ]) {
          final cell = decode(fixture('<c r="A1" s="STYLE"><v>-123</v></c>',
                  format: format))
              .tables
              .values
              .first
              .rows
              .first
              .first!;
          expect(cell.value, IntCellValue(-123), reason: format);
        }
        final date = decode(fixture('<c r="A1" s="STYLE"><v>46271</v></c>',
                format: 'M/D/YYYY'))
            .tables
            .values
            .first
            .rows
            .first
            .first!;
        expect(date.value, const DateCellValue(year: 2026, month: 9, day: 6));
        for (final format in ['[h]:mm:ss', '[M]', '[s]']) {
          expect(
              NumFormat.custom(formatCode: format), isA<DateTimeNumFormat>());
        }
      });
    });
  }
  test('1904 metadata and rich text survive save and re-decode', () {
    final workbook = Excel.decodeBytes(fixture(
        '<c r="A1" t="inlineStr"><is><t>A</t><r><t>B</t></r><t>C</t></is></c>',
        date1904: '1'));
    final restored = Excel.decodeBytes(workbook.encode()!);
    expect(restored.uses1904DateSystem, isTrue);
    expect(
        restored.tables.values.first.rows.first.first!.value.toString(), 'ABC');
  });
}

List<int> fixture(String cells,
    {String? shared, String? date1904, String? format}) {
  final workbook = Excel.createExcel();
  workbook.tables.values.first.updateCell(
      CellIndex.indexByString('A1'), IntCellValue(1),
      cellStyle:
          CellStyle(numberFormat: NumFormat.custom(formatCode: '0.0000')));
  final archive = ZipDecoder().decodeBytes(workbook.encode()!);
  final sheet = archive.files.firstWhere(
      (f) => f.name.startsWith('xl/worksheets/') && f.name.endsWith('.xml'));
  if (format != null) {
    final styles = XmlDocument.parse(
        utf8.decode(archive.findFile('xl/styles.xml')!.content));
    styles.findAllElements('numFmt').first.setAttribute('formatCode', format);
    archive.addFile(ArchiveFile.string('xl/styles.xml', styles.toXmlString()));
  }
  final original = XmlDocument.parse(utf8.decode(sheet.content));
  final style = original.findAllElements('c').first.getAttribute('s')!;
  archive.addFile(ArchiveFile.string(
      sheet.name,
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
      '<sheetData><row r="1">${cells.replaceAll('STYLE', style)}</row></sheetData></worksheet>'));
  if (shared != null) {
    archive.addFile(ArchiveFile.string('xl/sharedStrings.xml',
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si>$shared</si></sst>'));
  }
  if (date1904 != null) {
    final doc = XmlDocument.parse(
        utf8.decode(archive.findFile('xl/workbook.xml')!.content));
    final props = doc.findAllElements('workbookPr').firstOrNull;
    if (props != null) {
      props.setAttribute('date1904', date1904);
    } else {
      doc.rootElement.children.insert(
          0,
          XmlElement(XmlName('workbookPr'),
              [XmlAttribute(XmlName('date1904'), date1904)]));
    }
    archive.addFile(ArchiveFile.string('xl/workbook.xml', doc.toXmlString()));
  }
  return ZipEncoder().encode(archive);
}
