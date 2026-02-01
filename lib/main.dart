import 'dart:convert';
import 'dart:js' as js;
import 'dart:typed_data';

import 'package:archive/archive.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter/foundation.dart';
import 'package:flutter/material.dart';
import 'package:xml/xml.dart';

void main() => runApp(const MyApp());

const String kAppVersion = String.fromEnvironment('APP_VERSION', defaultValue: 'dev');

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Excel 矩阵压缩器',
      theme: ThemeData(
        colorScheme: ColorScheme.fromSeed(seedColor: Colors.blue),
        useMaterial3: true,
      ),
      home: const HomePage(),
      debugShowCheckedModeBanner: false,
    );
  }
}

class HomePage extends StatefulWidget {
  const HomePage({super.key});

  @override
  State<HomePage> createState() => _HomePageState();
}

class _HomePageState extends State<HomePage> {
  Uint8List? _fileBytes;
  String _status = '准备就绪';
  bool _busy = false;

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(title: const Text('Excel 智能压缩 (多线程版)'), centerTitle: true),
      body: Center(
        child: Container(
          constraints: const BoxConstraints(maxWidth: 600),
          padding: const EdgeInsets.all(24),
          child: Column(
            mainAxisAlignment: MainAxisAlignment.center,
            children: [
              _buildStatusCard(),
              const SizedBox(height: 32),
              if (_busy) const CircularProgressIndicator(),
              const SizedBox(height: 32),
              Row(
                mainAxisAlignment: MainAxisAlignment.center,
                children: [
                  ElevatedButton.icon(
                    onPressed: _busy ? null : _pickFile,
                    icon: const Icon(Icons.file_upload),
                    label: const Text('选择文件'),
                  ),
                  const SizedBox(width: 16),
                  ElevatedButton.icon(
                    onPressed: (_fileBytes == null || _busy) ? null : _filterAndExport,
                    icon: const Icon(Icons.bolt),
                    label: const Text('开始压缩导出'),
                    style: ElevatedButton.styleFrom(backgroundColor: Colors.blue[50]),
                  ),
                ],
              ),
              const SizedBox(height: 24),
              Text(
                '版本: $kAppVersion',
                style: const TextStyle(fontSize: 12, color: Colors.black45),
              ),
            ],
          ),
        ),
      ),
    );
  }

  Widget _buildStatusCard() {
    return Card(
      elevation: 0,
      color: _fileBytes == null ? Colors.grey[100] : Colors.blue[50],
      shape: RoundedRectangleBorder(
        borderRadius: BorderRadius.circular(12),
        side: BorderSide(color: Colors.blue[100]!),
      ),
      child: Padding(
        padding: const EdgeInsets.all(20),
        child: Column(
          children: [
            const Icon(Icons.description, size: 48, color: Colors.blue),
            const SizedBox(height: 12),
            Text(
              _status,
              textAlign: TextAlign.center,
              style: const TextStyle(fontSize: 14, fontWeight: FontWeight.w500),
            ),
          ],
        ),
      ),
    );
  }

  Future<void> _pickFile() async {
    try {
      final result = await FilePicker.platform.pickFiles(
        type: FileType.custom,
        allowedExtensions: ['xlsx'],
        allowMultiple: false,
        withData: true,
      );
      if (result != null && result.files.single.bytes != null) {
        setState(() {
          _fileBytes = result.files.single.bytes;
          _status = '已选中: ${result.files.single.name}';
        });
      }
    } catch (e) {
      setState(() => _status = '选取失败: $e');
    }
  }

  Future<void> _filterAndExport() async {
    if (_fileBytes == null) return;
    setState(() {
      _busy = true;
      _status = '正在读取文件并分配线程...';
    });

    try {
      final outputBytes = await compute(_heavyProcessTask, _fileBytes!);
      final filename = '压缩报表_${DateTime.now().millisecondsSinceEpoch}.xlsx';
      _downloadBytes(outputBytes, filename);
      setState(() => _status = '处理完成！请查看浏览器下载');
    } catch (e) {
      setState(() => _status = '处理失败: $e');
    } finally {
      setState(() => _busy = false);
    }
  }
}

void _downloadBytes(Uint8List bytes, String filename) {
  final b64 = base64Encode(bytes);
  js.context.callMethod('eval', [
    """
    var a = document.createElement('a');
    a.href = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,$b64';
    a.download = '$filename';
    a.click();
    """
  ]);
}

// --- 后台线程处理逻辑 (Top-level Functions) ---

Uint8List _heavyProcessTask(Uint8List bytes) {
  final archive = ZipDecoder().decodeBytes(bytes);
  final outArchive = Archive();

  final ssFile = _findInArchive(archive, 'xl/sharedStrings.xml');
  final sharedStrings = ssFile == null ? <String>[] : _parseStrings(ssFile.content as List<int>);
  final stylesFile = _findInArchive(archive, 'xl/styles.xml');
  XmlDocument? stylesDoc;
  _StyleManager? styleManager;
  if (stylesFile != null) {
    stylesDoc = XmlDocument.parse(utf8.decode(stylesFile.content as List<int>));
    styleManager = _StyleManager.fromDoc(stylesDoc);
  }

  for (final file in archive) {
    if (file.name.startsWith('xl/worksheets/sheet') && file.name.endsWith('.xml')) {
      final doc = XmlDocument.parse(utf8.decode(file.content as List<int>));
      final result = _filterCore(doc, 7, sharedStrings, styleManager);
      final encoded = utf8.encode(result.toXmlString(pretty: false));
      outArchive.addFile(ArchiveFile(file.name, encoded.length, encoded));
    } else if (file.name == 'xl/styles.xml' && stylesDoc != null) {
      final encoded = utf8.encode(stylesDoc.toXmlString(pretty: false));
      outArchive.addFile(ArchiveFile(file.name, encoded.length, encoded));
    } else {
      outArchive.addFile(ArchiveFile(file.name, file.size, file.content as List<int>));
    }
  }
  return Uint8List.fromList(ZipEncoder().encode(outArchive)!);
}

XmlDocument _filterCore(
  XmlDocument doc,
  int headerRows,
  List<String> sharedStrings,
  _StyleManager? styleManager,
) {
  final sheetData = doc.findAllElements('sheetData').first;
  final rows = sheetData.findElements('row').toList();

  final blackList = {'E', 'G', 'H', 'I', 'K', 'L', 'N', 'O', 'P', 'Q'};
  Set<String> keptColLetters = {'A', 'B', 'C'};

  for (var row in rows) {
    int rIdx = int.tryParse(row.getAttribute('r') ?? '') ?? 0;
    if (rIdx <= headerRows) continue;
    for (var cell in row.findElements('c')) {
      String col = (cell.getAttribute('r') ?? '').replaceAll(RegExp(r'[0-9]'), '');
      if (!blackList.contains(col) && !keptColLetters.contains(col)) {
        if (_isCellValid(cell, sharedStrings)) keptColLetters.add(col);
      }
    }
  }

  List<String> sortedOldCols = keptColLetters.toList()
    ..sort((a, b) => _colToIdx(a).compareTo(_colToIdx(b)));

  int maxRow = 0;
  for (var r in rows) {
    int idx = int.tryParse(r.getAttribute('r') ?? '') ?? 0;
    if (idx > maxRow) maxRow = idx;
  }

  List<XmlElement> finalRows = [];
  int nextRowIdx = 1;
  final Map<int, int> rowMap = {};

  for (var row in rows) {
    int oldR = int.tryParse(row.getAttribute('r') ?? '') ?? 0;
    if (oldR == 3 || oldR == 4) {
      continue;
    }
    bool isFixed = oldR <= headerRows || oldR == maxRow;
    bool hasData = false;

    if (!isFixed) {
      for (var c in row.findElements('c')) {
        final col = (c.getAttribute('r') ?? '').replaceAll(RegExp(r'[0-9]'), '');
        if (keptColLetters.contains(col) && _isCellValid(c, sharedStrings)) {
          hasData = true;
          break;
        }
      }
    }

    if (isFixed || hasData) {
      final newRow = _cloneElement(row);
      newRow.setAttribute('r', nextRowIdx.toString());
      rowMap[oldR] = nextRowIdx;
      List<XmlElement> newCells = [];
      for (int i = 0; i < sortedOldCols.length; i++) {
        String oldCol = sortedOldCols[i];
        String newCol = _idxToCol(i);
        var cell = row.findElements('c').firstWhere(
              (c) => (c.getAttribute('r') ?? '').startsWith(oldCol),
              orElse: () => XmlElement(XmlName('null')),
            );
        if (cell.name.local == 'c') {
          final newCell = _cloneElement(cell);
          newCell.setAttribute('r', '$newCol$nextRowIdx');
          _applyAccountingFormatIfZero(newCell, sharedStrings, styleManager);
          newCells.add(newCell);
        }
      }
      _replaceChildren(newRow, newCells);
      finalRows.add(newRow);
      nextRowIdx++;
    }
  }

  _replaceChildren(sheetData, finalRows);

  List<String> merges = [];

  doc.findAllElements('mergeCell').forEach((m) {
    final ref = (m.getAttribute('ref') ?? '').split(':');
    if (ref.length == 2) {
      int r1 = int.tryParse(ref[0].replaceAll(RegExp(r'[^0-9]'), '')) ?? 0;
      int r2 = int.tryParse(ref[1].replaceAll(RegExp(r'[^0-9]'), '')) ?? 0;
      if (r1 <= headerRows && r2 <= headerRows) {
        final newR1 = rowMap[r1];
        final newR2 = rowMap[r2];
        if (newR1 == null || newR2 == null) {
          return;
        }
        String c1 = ref[0].replaceAll(RegExp(r'[0-9]'), '');
        String c2 = ref[1].replaceAll(RegExp(r'[0-9]'), '');
        int startI = -1;
        int endI = -1;
        for (int i = 0; i < sortedOldCols.length; i++) {
          if (_colToIdx(sortedOldCols[i]) >= _colToIdx(c1) && startI == -1) startI = i;
          if (_colToIdx(sortedOldCols[i]) <= _colToIdx(c2)) endI = i;
        }
        if (startI != -1 && endI >= startI) {
          merges.add('${_idxToCol(startI)}$newR1:${_idxToCol(endI)}$newR2');
        }
      }
    }
  });

  String? lastA;
  int start = -1;
  for (int i = headerRows; i < finalRows.length; i++) {
    String cur = _getVal(finalRows[i], 'A', sharedStrings);
    int rowNum = i + 1;
    if (cur.isNotEmpty && cur == lastA) {
    } else {
      if (start != -1 && (rowNum - 1) > start) merges.add("A$start:A${rowNum - 1}");
      lastA = cur;
      start = rowNum;
    }
    if (i == finalRows.length - 1 && start != -1 && rowNum > start) {
      merges.add("A$start:A$rowNum");
    }
  }

  doc.findAllElements('mergeCells').forEach((e) => e.parent?.children.remove(e));
  if (merges.isNotEmpty) {
    final builder = XmlBuilder();
    builder.element('mergeCells', attributes: {'count': merges.length.toString()}, nest: () {
      for (var r in merges) builder.element('mergeCell', attributes: {'ref': r});
    });
    doc.rootElement.children.add(builder.buildFragment());
  }

  return doc;
}

void _applyAccountingFormatIfZero(
  XmlElement cell,
  List<String> sharedStrings,
  _StyleManager? styleManager,
) {
  if (styleManager == null) return;
  final t = cell.getAttribute('t');
  final parsed = _parseCellNumericValue(cell, t, sharedStrings);
  if (parsed == null) return;
  final n = parsed.$1;
  final wasString = parsed.$2;
  if (n == null || n.abs() > 0.000001) return;
  if (wasString) {
    _convertCellToNumericZero(cell);
  }
  final baseStyleIndex = int.tryParse(cell.getAttribute('s') ?? '') ?? 0;
  final accountingStyleIndex = styleManager.accountingStyleFor(baseStyleIndex);
  cell.setAttribute('s', accountingStyleIndex.toString());
}

(double?, bool)? _parseCellNumericValue(
  XmlElement cell,
  String? t,
  List<String> sharedStrings,
) {
  if (t == 's') {
    final v = cell.getElement('v')?.text;
    if (v == null) return null;
    final idx = int.tryParse(v);
    if (idx == null || idx < 0 || idx >= sharedStrings.length) return null;
    return (double.tryParse(sharedStrings[idx].trim()), true);
  }
  if (t == 'inlineStr') {
    final text = cell
        .findElements('is')
        .expand((e) => e.findElements('t'))
        .map((e) => e.text)
        .join();
    if (text.isEmpty) return null;
    return (double.tryParse(text.trim()), true);
  }
  final v = cell.getElement('v')?.text;
  if (v == null) return null;
  if (t == 'str') {
    return (double.tryParse(v.trim()), true);
  }
  return (double.tryParse(v.trim()), false);
}

void _convertCellToNumericZero(XmlElement cell) {
  cell.attributes.removeWhere((a) => a.name.local == 't');
  cell.children.removeWhere((node) => node is XmlElement && node.name.local == 'is');
  var value = cell.getElement('v');
  if (value == null) {
    value = XmlElement(_StyleManager._nsName(cell, 'v'));
    cell.children.add(value);
  }
  value.children
    ..clear()
    ..add(XmlText('0'));
}

bool _isCellValid(XmlElement cell, List<String> ss) {
  final v = cell.getElement('v')?.text;
  if (v == null) return false;
  String content = cell.getAttribute('t') == 's' ? (int.tryParse(v) != null ? ss[int.parse(v)] : '') : v;
  final n = double.tryParse(content.trim());
  return n != null && n.abs() > 0.000001;
}

String _getVal(XmlElement row, String col, List<String> ss) {
  for (var c in row.findElements('c')) {
    if ((c.getAttribute('r') ?? '').startsWith(col)) {
      final v = c.getElement('v')?.text;
      if (v == null) return '';
      final t = c.getAttribute('t');
      return t == 's' ? (int.tryParse(v) != null ? ss[int.parse(v)] : '') : v;
    }
  }
  return '';
}

List<String> _parseStrings(List<int> bytes) {
  final doc = XmlDocument.parse(utf8.decode(bytes));
  return doc.findAllElements('si').map((si) => si.findAllElements('t').map((e) => e.text).join()).toList();
}

ArchiveFile? _findInArchive(Archive a, String n) {
  for (var f in a) {
    if (f.name == n) return f;
  }
  return null;
}

int _colToIdx(String c) {
  int idx = 0;
  for (int i = 0; i < c.length; i++) {
    idx = idx * 26 + (c.codeUnitAt(i) - 64);
  }
  return idx;
}

String _idxToCol(int i) {
  String c = '';
  i++;
  while (i > 0) {
    int r = (i - 1) % 26;
    c = String.fromCharCode(65 + r) + c;
    i = (i - r) ~/ 26;
  }
  return c;
}

class _StyleManager {
  static const String _accountingFormatCode =
      '_,* #,##0.00_);_,* (#,##0.00);_,* "-"??_);_(@_)';

  final XmlElement cellXfs;
  final List<XmlElement> cellXfList;
  final int accountingNumFmtId;
  final Map<int, int> _accountingStyleCache = {};

  _StyleManager._(
    this.cellXfs,
    this.cellXfList,
    this.accountingNumFmtId,
  );

  factory _StyleManager.fromDoc(XmlDocument doc) {
    final styleSheet = doc.rootElement;
    final numFmts = _ensureNumFmts(styleSheet);
    final accountingNumFmtId = _ensureAccountingNumFmt(numFmts);
    final cellXfs = _ensureCellXfs(styleSheet);
    final cellXfList = cellXfs.findElements('xf').toList();
    if (cellXfList.isEmpty) {
      final defaultXf = XmlElement(
        _nsName(cellXfs, 'xf'),
        [
          XmlAttribute(XmlName('numFmtId'), '0'),
          XmlAttribute(XmlName('fontId'), '0'),
          XmlAttribute(XmlName('fillId'), '0'),
          XmlAttribute(XmlName('borderId'), '0'),
          XmlAttribute(XmlName('xfId'), '0'),
        ],
      );
      cellXfs.children.add(defaultXf);
      cellXfList.add(defaultXf);
      cellXfs.setAttribute('count', cellXfList.length.toString());
    }
    return _StyleManager._(cellXfs, cellXfList, accountingNumFmtId);
  }

  int accountingStyleFor(int baseStyleIndex) {
    final normalized = (baseStyleIndex >= 0 && baseStyleIndex < cellXfList.length)
        ? baseStyleIndex
        : 0;
    final cached = _accountingStyleCache[normalized];
    if (cached != null) return cached;

  final baseXf = cellXfList[normalized];
  final newXf = _cloneElement(baseXf);
    newXf.setAttribute('numFmtId', accountingNumFmtId.toString());
    newXf.setAttribute('applyNumberFormat', '1');
    cellXfs.children.add(newXf);
    cellXfList.add(newXf);
    final newIndex = cellXfList.length - 1;
    cellXfs.setAttribute('count', cellXfList.length.toString());
    _accountingStyleCache[normalized] = newIndex;
    return newIndex;
  }

  static XmlElement _ensureNumFmts(XmlElement styleSheet) {
    final existing = styleSheet.getElement('numFmts');
    if (existing != null) return existing;
    final nsName = _nsName(styleSheet, 'numFmts');
    final numFmts = XmlElement(
      nsName,
      [XmlAttribute(XmlName('count'), '0')],
    );
    final children = styleSheet.children;
    int insertIndex = children.indexWhere(
      (node) => node is XmlElement && node.name.local == 'fonts',
    );
    if (insertIndex == -1) insertIndex = children.length;
    children.insert(insertIndex, numFmts);
    return numFmts;
  }

  static int _ensureAccountingNumFmt(XmlElement numFmts) {
    int maxId = 163;
    for (final numFmt in numFmts.findElements('numFmt')) {
      final formatCode = numFmt.getAttribute('formatCode');
      final id = int.tryParse(numFmt.getAttribute('numFmtId') ?? '') ?? 0;
      if (id > maxId) maxId = id;
      if (formatCode == _accountingFormatCode) {
        return id;
      }
    }
    final newId = maxId + 1;
    final newNumFmt = XmlElement(
      _nsName(numFmts, 'numFmt'),
      [
        XmlAttribute(XmlName('numFmtId'), newId.toString()),
        XmlAttribute(XmlName('formatCode'), _accountingFormatCode),
      ],
    );
    numFmts.children.add(newNumFmt);
    numFmts.setAttribute('count', numFmts.findElements('numFmt').length.toString());
    return newId;
  }

  static XmlElement _ensureCellXfs(XmlElement styleSheet) {
    final existing = styleSheet.getElement('cellXfs');
    if (existing != null) return existing;
    final nsName = _nsName(styleSheet, 'cellXfs');
    final cellXfs = XmlElement(
      nsName,
      [XmlAttribute(XmlName('count'), '1')],
      [
        XmlElement(
          _nsName(styleSheet, 'xf'),
          [
            XmlAttribute(XmlName('numFmtId'), '0'),
            XmlAttribute(XmlName('fontId'), '0'),
            XmlAttribute(XmlName('fillId'), '0'),
            XmlAttribute(XmlName('borderId'), '0'),
            XmlAttribute(XmlName('xfId'), '0'),
          ],
        ),
      ],
    );
    final children = styleSheet.children;
    int insertIndex = children.indexWhere(
      (node) => node is XmlElement && node.name.local == 'cellStyles',
    );
    if (insertIndex == -1) insertIndex = children.length;
    children.insert(insertIndex, cellXfs);
    return cellXfs;
  }

  static XmlName _nsName(XmlElement parent, String local) {
    return XmlName(local, parent.name.prefix);
  }
}

XmlElement _cloneElement(XmlElement element) {
  return element.copy();
}

void _replaceChildren(XmlElement parent, List<XmlElement> nodes) {
  parent.children.clear();
  for (final node in nodes) {
    parent.children.add(node.hasParent ? node.copy() : node);
  }
}
