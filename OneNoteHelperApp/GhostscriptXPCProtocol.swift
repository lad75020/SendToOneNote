import Foundation

/// XPC interface for PostScript -> PDF conversion.
@objc public protocol GhostscriptXPCProtocol {
    /// Convert PostScript at `psPath` to PDF at `pdfPath`.
    /// - Returns: (ok, logs) where logs is combined stdout/stderr.
    func convertPS(psPath: String, pdfPath: String, reply: @escaping (Bool, String) -> Void)
}
