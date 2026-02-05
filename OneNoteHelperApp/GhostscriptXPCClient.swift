import Foundation

/// Thin client used by the sandboxed app to talk to the embedded XPC service.
///
/// Note: we keep it `final` + static for simplicity.
final class GhostscriptXPCClient {
    /// Must match the XPC service bundle identifier.
    static let serviceName = "fr.dubertrand.OneNoteGhostscriptXPC"

    private static func embeddedServicePathHint() -> String {
        // When embedded, the service should live here:
        //   <App>.app/Contents/XPCServices/<Service>.xpc
        let appURL = Bundle.main.bundleURL
        let xpcDir = appURL.appendingPathComponent("Contents/XPCServices", isDirectory: true)
        let xpcURL = xpcDir.appendingPathComponent("OneNoteGhostscriptXPC.xpc", isDirectory: true)
        return xpcURL.path
    }

    static func convertPS(psPath: String, pdfPath: String, timeoutSeconds: TimeInterval = 60) -> (ok: Bool, logs: String) {
        let sem = DispatchSemaphore(value: 0)
        var result: (Bool, String) = (false, "")

        let conn = NSXPCConnection(serviceName: serviceName)
        conn.remoteObjectInterface = NSXPCInterface(with: GhostscriptXPCProtocol.self)
        conn.resume()

        let proxy = conn.remoteObjectProxyWithErrorHandler { err in
            // This error is often caused by the XPC not being embedded/codesigned correctly.
            result = (false, "XPC error: \(err.localizedDescription). Expected embedded service at: \(embeddedServicePathHint())")
            sem.signal()
        } as? GhostscriptXPCProtocol

        guard let proxy else {
            conn.invalidate()
            return (false, "XPC: failed to get proxy. Expected embedded service at: \(embeddedServicePathHint())")
        }

        proxy.convertPS(psPath: psPath, pdfPath: pdfPath) { ok, logs in
            result = (ok, logs)
            sem.signal()
        }

        let waitRes = sem.wait(timeout: .now() + timeoutSeconds)
        conn.invalidate()

        if waitRes == .timedOut {
            return (false, "XPC: timed out after \(timeoutSeconds)s. Expected embedded service at: \(embeddedServicePathHint())")
        }

        return result
    }
}
