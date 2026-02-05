import Foundation

final class GhostscriptXPCService: NSObject, GhostscriptXPCProtocol {
    func convertPS(psPath: String, pdfPath: String, reply: @escaping (Bool, String) -> Void) {
        let fm = FileManager.default

        guard let exeURL = Bundle.main.executableURL else {
            reply(false, "XPC: Bundle.main.executableURL missing")
            return
        }

        // gs is placed next to the XPC executable: <Service>.xpc/Contents/MacOS/gs
        let gsURL = exeURL.deletingLastPathComponent().appendingPathComponent("gs")

        guard fm.isExecutableFile(atPath: gsURL.path) else {
            reply(false, "XPC: gs missing/not executable at \(gsURL.path)")
            return
        }

        let proc = Process()
        proc.executableURL = gsURL
        proc.arguments = [
            "-dSAFER",
            "-dBATCH",
            "-dNOPAUSE",
            "-sDEVICE=pdfwrite",
            "-dCompatibilityLevel=1.4",
            "-sOutputFile=\(pdfPath)",
            psPath
        ]

        let outPipe = Pipe()
        let errPipe = Pipe()
        proc.standardOutput = outPipe
        proc.standardError = errPipe

        do {
            try proc.run()
            proc.waitUntilExit()
        } catch {
            reply(false, "XPC: failed to start gs: \(error.localizedDescription)")
            return
        }

        let out = String(data: outPipe.fileHandleForReading.readDataToEndOfFile(), encoding: .utf8) ?? ""
        let err = String(data: errPipe.fileHandleForReading.readDataToEndOfFile(), encoding: .utf8) ?? ""
        let logs = ([out, err].joined(separator: "\n")).trimmingCharacters(in: .whitespacesAndNewlines)

        if proc.terminationReason == .uncaughtSignal {
            reply(false, "XPC: gs killed by signal \(proc.terminationStatus). Logs: \(logs)")
            return
        }

        if proc.terminationStatus != 0 {
            reply(false, "XPC: gs exit=\(proc.terminationStatus). Logs: \(logs)")
            return
        }

        reply(true, logs)
    }
}

final class ServiceDelegate: NSObject, NSXPCListenerDelegate {
    private let exported = GhostscriptXPCService()

    func listener(_ listener: NSXPCListener, shouldAcceptNewConnection newConnection: NSXPCConnection) -> Bool {
        newConnection.exportedInterface = NSXPCInterface(with: GhostscriptXPCProtocol.self)
        newConnection.exportedObject = exported
        newConnection.resume()
        return true
    }
}

let delegate = ServiceDelegate()
let listener = NSXPCListener.service()
listener.delegate = delegate
listener.resume()
