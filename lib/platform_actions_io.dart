import 'dart:io';
import 'dart:typed_data';

import 'package:file_picker/file_picker.dart';

Future<void> platformDownloadBytes(Uint8List bytes, String filename) async {
  final path = await FilePicker.platform.saveFile(
    dialogTitle: '保存文件',
    fileName: filename,
    type: FileType.custom,
    allowedExtensions: const ['xlsx'],
  );
  if (path == null || path.isEmpty) return;
  final file = File(path);
  await file.writeAsBytes(bytes, flush: true);
}

Future<void> platformClearWebCacheAndReload() async {}
