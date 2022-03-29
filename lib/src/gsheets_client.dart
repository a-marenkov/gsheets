import 'package:googleapis_auth/auth_io.dart';

import 'gsheets.dart';

abstract class GSheetsAuth {
  const GSheetsAuth._();

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

    if (credentials != null) {
      return clientViaServiceAccount(credentials, scopes);
    }

    // if (clientId != null) {
    //   final flow = await createImplicitBrowserFlow(clientId, scopes);
    //   final client = await flow.clientViaUserConsent();
    //   flow.close();
    //   return client;
    // }

    throw GSheetsException('clientId, credentials or client must be provided');
  }
}
