import 'dart:typed_data';

import 'platform_actions_stub.dart'
    if (dart.library.html) 'platform_actions_web.dart'
    if (dart.library.io) 'platform_actions_io.dart';

Future<void> downloadBytes(Uint8List bytes, String filename) =>
    platformDownloadBytes(bytes, filename);

Future<void> clearWebCacheAndReload() => platformClearWebCacheAndReload();
