import 'dart:convert';
import 'dart:io';

import 'package:archive/archive.dart';
import 'package:excel/excel.dart';
import 'package:test/test.dart';
import 'package:xml/xml.dart';

void main() {
  group('absolute workbook relationship targets', () {
    final sourceBytes =
        File('test/test_resources/example.xlsx').readAsBytesSync();

    test('keeps relative relationship targets working', () {
      final excel = Excel.decodeBytes(sourceBytes);

      expect(_washingtonValue(excel), 'Washington');
    });

    test('decodes an absolute worksheet target from bytes', () {
      final bytes = _makeWorkbookRelationshipTargetsAbsolute(
        sourceBytes,
        {'worksheet'},
      );

      final excel = Excel.decodeBytes(bytes);

      expect(_washingtonValue(excel), 'Washington');
    });

    test('decodes absolute internal targets from a buffer', () {
      final bytes = _makeWorkbookRelationshipTargetsAbsolute(
        sourceBytes,
        {'worksheet', 'styles', 'sharedStrings'},
      );

      final excel = Excel.decodeBuffer(InputMemoryStream(bytes));

      expect(_washingtonValue(excel), 'Washington');
    });

    test('writes normalized internal targets and supports round-trip', () {
      final bytes = _makeWorkbookRelationshipTargetsAbsolute(
        sourceBytes,
        {'worksheet', 'styles', 'sharedStrings'},
      );
      final excel = Excel.decodeBytes(bytes);

      final encoded = excel.encode()!;
      final targets = _workbookRelationshipTargets(encoded);
      final decoded = Excel.decodeBytes(encoded);

      expect(targets['worksheet'], 'worksheets/sheet1.xml');
      expect(targets['styles'], 'styles.xml');
      expect(targets['sharedStrings'], 'sharedStrings.xml');
      expect(_washingtonValue(decoded), 'Washington');
    });
  });
}

String _washingtonValue(Excel excel) {
  return excel.tables['Sheet1']!.rows[1][1]!.value.toString();
}

List<int> _makeWorkbookRelationshipTargetsAbsolute(
  List<int> bytes,
  Set<String> relationshipNames,
) {
  const relationshipsPath = 'xl/_rels/workbook.xml.rels';
  final archive = ZipDecoder().decodeBytes(bytes);
  final relationshipsFile = archive.findFile(relationshipsPath)!;
  final document = XmlDocument.parse(utf8.decode(relationshipsFile.content));

  for (final relationship in document.findAllElements('Relationship')) {
    final type = relationship.getAttribute('Type');
    final relationshipName = type?.split('/').last;
    final target = relationship.getAttribute('Target');
    if (relationshipName != null &&
        relationshipNames.contains(relationshipName) &&
        target != null &&
        !target.startsWith('/')) {
      relationship.setAttribute('Target', '/xl/$target');
    }
  }

  archive.addFile(
    ArchiveFile.string(relationshipsPath, document.toXmlString()),
  );
  return ZipEncoder().encode(archive);
}

Map<String, String> _workbookRelationshipTargets(List<int> bytes) {
  final archive = ZipDecoder().decodeBytes(bytes);
  final relationshipsFile = archive.findFile('xl/_rels/workbook.xml.rels')!;
  final document = XmlDocument.parse(utf8.decode(relationshipsFile.content));

  return {
    for (final relationship in document.findAllElements('Relationship'))
      if (relationship.getAttribute('Type') case final String type)
        type.split('/').last: relationship.getAttribute('Target')!,
  };
}
