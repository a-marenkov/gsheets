import 'package:googleapis_auth/auth_browser.dart';
import 'package:googleapis_auth/auth_io.dart';
import 'package:googleapis_auth/googleapis_auth.dart';
import 'package:gsheets/gsheets.dart';

abstract class GSheetsClient {
  const GSheetsClient._();

  static Future<AutoRefreshingAuthClient> auth({
    required final Future<AutoRefreshingAuthClient>? client,
    required final List<String>? scopes,
    required final ClientId? clientId,
    required final ServiceAccountCredentials? credentials,
  }) async {
    assert(
      clientId != null || credentials != null || client != null,
      'clientId, credentials or client must be provided',
    );

    if (client != null) {
      return client;
    }

    if (scopes == null) {
      throw GSheetsException('scopes must be provided');
    }

    if (clientId != null) {
      final flow = await createImplicitBrowserFlow(clientId, scopes);
      final client = await flow.clientViaUserConsent();
      flow.close();
      return client;
    }

    if (credentials != null) {
      return clientViaServiceAccount(credentials, scopes);
    }

    throw GSheetsException('clientId, credentials or client must be provided');
  }
}
