import 'dart:html' as html;
import 'dart:typed_data';

Future<void> platformDownloadBytes(Uint8List bytes, String filename) async {
  final blob = html.Blob(
    [bytes],
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  );
  final url = html.Url.createObjectUrlFromBlob(blob);
  final anchor = html.AnchorElement(href: url)
    ..download = filename
    ..style.display = 'none';
  html.document.body?.children.add(anchor);
  anchor.click();
  anchor.remove();
  html.Url.revokeObjectUrl(url);
}

Future<void> platformClearWebCacheAndReload() async {
  Future<void> finish() async {
    html.window.location.reload();
  }

  try {
    final registrations = await html.window.navigator.serviceWorker?.getRegistrations();
    if (registrations != null) {
      for (final registration in registrations) {
        await registration.unregister();
      }
    }
    if (html.window.caches != null) {
      final keys = await html.window.caches!.keys();
      for (final key in keys) {
        await html.window.caches!.delete(key);
      }
    }
    await finish();
  } catch (_) {
    await finish();
  }
}
