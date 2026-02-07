import SwiftUI
import AppKit
import CoreGraphics
import PDFKit
import ImageIO
import UniformTypeIdentifiers
import os.log
import MSAL
import Foundation
import Darwin

private let OneNoteHelperWatcherQueue = DispatchQueue(label: "fr.dubertrand.OneNoteHelperApp.watcher")

@MainActor
final class AppDelegate: NSObject, NSApplicationDelegate {
    static weak var shared: AppDelegate?

    private let graphScopes = ["Notes.ReadWrite", "User.Read"]

    // Graph selection storage
    private let targetSectionIdKey = "TargetSectionId"
    private let targetPageIdKey = "TargetPageId"

    private var dirSource: DispatchSourceFileSystemObject?
    private var pollTimer: DispatchSourceTimer?

    // Scan throttling (prevents event storms + re-entrant scans)
    private var scanWorkItem: DispatchWorkItem?
    private var isScanning: Bool = false

    private var interactiveAuthInProgress = false
    private var interactiveAuthWaiters: [(String?) -> Void] = []

    nonisolated private func log(_ message: String) {
        os_log("%{public}@", message)

        // Also log to a plain file for debugging when unified logging/UI logs are not visible.
        // Write inside the watch folder root so sandboxed builds can access it via security-scoped bookmark.
        let line = "[\(Date())] \(message)\n"
        if let data = line.data(using: .utf8) {
            let root = UserDefaults.standard.string(forKey: "WatchFolderPath") ?? "/Users/Shared/OneNoteHelper"
            let trimmed = root.trimmingCharacters(in: .whitespacesAndNewlines)
            let rootPath = trimmed.isEmpty ? "/Users/Shared/OneNoteHelper" : trimmed
            let url = URL(fileURLWithPath: rootPath, isDirectory: true).appendingPathComponent("helper.log")
            if FileManager.default.fileExists(atPath: url.path) {
                if let fh = try? FileHandle(forWritingTo: url) {
                    try? fh.seekToEnd()
                    try? fh.write(contentsOf: data)
                    try? fh.close()
                }
            } else {
                try? data.write(to: url, options: .atomic)
                try? FileManager.default.setAttributes([.posixPermissions: 0o666], ofItemAtPath: url.path)
            }
        }

        Task { @MainActor in
            LogStore.shared.append(message)
        }
    }

    private let fm = FileManager.default
    // Removed: private let watcherQueue = DispatchQueue(label: "fr.dubertrand.OneNoteHelperApp.watcher")

    private var msalApplication: MSALPublicClientApplication? {
        // Don't cache this: if Info.plist/UserDefaults change, we want to pick up
        // the new values after a rebuild / defaults write.
        guard let clientId = msalClientId, !clientId.isEmpty else {
            log("MSAL client ID missing. Set MSALClientId in Info.plist or UserDefaults.")
            return nil
        }

        let bundleId = Bundle.main.bundleIdentifier ?? ""
        let defaultRedirect = "msauth.\(bundleId)://auth"
        let authorityURLString = (msalAuthority?.trimmingCharacters(in: .whitespacesAndNewlines)).flatMap { $0.isEmpty ? nil : $0 } ?? "https://login.microsoftonline.com/organizations"
        let authority: MSALAuthority?
        if let url = URL(string: authorityURLString) {
            authority = try? MSALAADAuthority(url: url)
        } else {
            authority = nil
        }

        // MSAL config log removed (too noisy)
        do {
            let config = MSALPublicClientApplicationConfig(clientId: clientId, redirectUri: defaultRedirect, authority: authority)

            // Disable keychain sharing and use the app-private keychain group.
            // This avoids entitlement/keychain-group mismatches that manifest as OSStatus -34018.
            config.cacheConfig.keychainSharingGroup = Bundle.main.bundleIdentifier ?? "fr.dubertrand.OneNoteHelperApp"

            let app = try MSALPublicClientApplication(configuration: config)
            // MSAL configured logs removed (too noisy)
            return app
        } catch {
            let nsError = error as NSError
            log("MSAL configuration failed: domain=\(nsError.domain) code=\(nsError.code) userInfo=\(nsError.userInfo)")
            log("Ensure Info.plist defines a URL type with scheme 'msauth.\(bundleId)' and that the same redirect URI is registered in your Azure app.")
            return nil
        }
    }

    nonisolated private var graphEndpointBase: String {
        let raw = (Bundle.main.object(forInfoDictionaryKey: "GraphEndpoint") as? String)
            ?? UserDefaults.standard.string(forKey: "GraphEndpoint")
            ?? "https://graph.microsoft.com"
        let trimmed = raw.trimmingCharacters(in: .whitespacesAndNewlines)
        return trimmed.hasSuffix("/") ? String(trimmed.dropLast()) : trimmed
    }

    nonisolated private func graphURL(_ path: String) -> String {
        let p = path.hasPrefix("/") ? String(path.dropFirst()) : path
        return graphEndpointBase + "/v1.0/" + p
    }

    private func isURLSchemeRegistered(_ scheme: String) -> Bool {
        guard let urlTypes = Bundle.main.object(forInfoDictionaryKey: "CFBundleURLTypes") as? [[String: Any]] else { return false }
        for type in urlTypes {
            if let schemes = type["CFBundleURLSchemes"] as? [String], schemes.contains(scheme) {
                return true
            }
        }
        return false
    }

    private var msalClientId: String? {
        let v = (
            (Bundle.main.object(forInfoDictionaryKey: "MSALClientId") as? String)
            ?? UserDefaults.standard.string(forKey: "MSALClientId")
            ?? ""
        ).trimmingCharacters(in: .whitespacesAndNewlines)
        return v.isEmpty ? nil : v
    }

    private var msalRedirectUri: String? {
        let v = (
            (Bundle.main.object(forInfoDictionaryKey: "MSALRedirectUri") as? String)
            ?? UserDefaults.standard.string(forKey: "MSALRedirectUri")
            ?? ""
        ).trimmingCharacters(in: .whitespacesAndNewlines)
        return v.isEmpty ? nil : v
    }

    private var msalAuthority: String? {
        (Bundle.main.object(forInfoDictionaryKey: "MSALAuthority") as? String)
            ?? UserDefaults.standard.string(forKey: "MSALAuthority")
    }

    // MSAL logging removed (too noisy for daily use).
    
    nonisolated private func keychainSanityCheck() {
        let account = "OneNoteHelperKeychainTest"
        let service = "fr.dubertrand.OneNoteHelperApp"
        let testData = "ok".data(using: .utf8)!

        // Clean any existing item
        let query: [String: Any] = [kSecClass as String: kSecClassGenericPassword,
                                    kSecAttrAccount as String: account,
                                    kSecAttrService as String: service]
        SecItemDelete(query as CFDictionary)

        // Try to add
        var addQuery = query
        addQuery[kSecValueData as String] = testData
        let addStatus = SecItemAdd(addQuery as CFDictionary, nil)
        if addStatus != errSecSuccess {
            self.log("Keychain sanity check: add failed with status=\(addStatus)")
            return
        }

        // Try to read
        var readQuery = query
        readQuery[kSecReturnData as String] = kCFBooleanTrue
        var result: CFTypeRef?
        let copyStatus = SecItemCopyMatching(readQuery as CFDictionary, &result)
        if copyStatus != errSecSuccess {
            self.log("Keychain sanity check: read failed with status=\(copyStatus)")
        } else {
            self.log("Keychain sanity check: success")
        }

        // Cleanup
        SecItemDelete(query as CFDictionary)
    }

    func applicationDidFinishLaunching(_ notification: Notification) {
        AppDelegate.shared = self
        OneNoteTargetStore.shared.register(appDelegate: self)
        resolveAndStartSecurityScopedAccessIfNeeded()
        startWatchingIncomingFolder()
        Task { @MainActor in
            OneNoteTargetStore.shared.refreshAll()
            self.keychainSanityCheck()
        }
    }

    func application(_ application: NSApplication, open urls: [URL]) {
        log("application:openURLs count=\(urls.count)")
        for url in urls {
            log("application:openURLs url=\(url.absoluteString)")
            handle(url: url)
        }
    }

    private func handle(url: URL) {
        guard url.scheme == "onenote-helper", url.host == "import" else { return }
        log("Received URL: \(url.absoluteString)")
        guard let components = URLComponents(url: url, resolvingAgainstBaseURL: false),
              let queryItems = components.queryItems else { return }

        func value(_ name: String) -> String? { queryItems.first { $0.name == name }?.value }

        guard let file = value("file") else { return }
        let title = value("title") ?? "Printed Document"
        let user = value("user") ?? ""
        let job = value("job") ?? ""

        importFile(filePath: file, title: title, user: user, job: job, completion: nil)
    }

    private func importFile(filePath: String, title: String, user: String, job: String, completion: ((Bool) -> Void)?) {
        log("Preparing upload: file=\(filePath) title=\(title) user=\(user) job=\(job)")
        LogStore.shared.activityState = .processing

        let showUIOnImport = UserDefaults.standard.bool(forKey: "ShowUIOnImport")
        if showUIOnImport {
            self.log("Import: showing UI for import")
            if NSApp.activationPolicy() != .regular {
                _ = NSApp.setActivationPolicy(.regular)
            }
            NSApp.activate(ignoringOtherApps: true)
            NSApp.windows.forEach { $0.makeKeyAndOrderFront(nil) }
        } else {
            self.log("Import: running headless (accessory)")
            if NSApp.activationPolicy() != .accessory {
                _ = NSApp.setActivationPolicy(.accessory)
            }
        }

        acquireGraphToken { token in
            guard let token else {
                LogStore.shared.activityState = .waiting
                completion?(false)
                return
            }
            self.log("Graph token acquired (len=\(token.count))")
            LogStore.shared.activityState = .uploading
            self.uploadSinglePage(token: token, filePath: filePath, title: title, user: user, job: job) { ok in
                LogStore.shared.activityState = .waiting
                completion?(ok)
            }
        }
    }

    private var securityScopedRootURL: URL?

    nonisolated private func helperRootDir() -> URL {
        // Root folder for the file-queue: must contain Incoming/Processing/Done/Failed.
        // Default matches the CUPS backend drop location.
        let raw = UserDefaults.standard.string(forKey: "WatchFolderPath") ?? "/Users/Shared/OneNoteHelper"
        let trimmed = raw.trimmingCharacters(in: .whitespacesAndNewlines)
        let path = trimmed.isEmpty ? "/Users/Shared/OneNoteHelper" : trimmed
        return URL(fileURLWithPath: path, isDirectory: true)
    }

    private func resolveAndStartSecurityScopedAccessIfNeeded() {
        // If we have a bookmark, resolve it and start accessing.
        guard let data = UserDefaults.standard.data(forKey: "WatchFolderBookmark") else { return }

        var isStale = false
        if let url = try? URL(resolvingBookmarkData: data,
                              options: [.withSecurityScope],
                              relativeTo: nil,
                              bookmarkDataIsStale: &isStale) {
            if isStale {
                self.log("Watch folder bookmark is stale; please re-select the folder in UI")
            }
            securityScopedRootURL = url
            if url.startAccessingSecurityScopedResource() {
                self.log("Security-scoped access granted for watch folder: \(url.path)")
            } else {
                self.log("Security-scoped access FAILED for watch folder: \(url.path)")
            }
        }
    }

    private func stopSecurityScopedAccessIfNeeded() {
        securityScopedRootURL?.stopAccessingSecurityScopedResource()
        securityScopedRootURL = nil
    }

    nonisolated private func folderURL(_ name: String) -> URL {
        helperRootDir().appendingPathComponent(name, isDirectory: true)
    }

    nonisolated private func ensureFolders() {
        let dirs = [folderURL("Incoming"), folderURL("Processing"), folderURL("Done"), folderURL("Failed")]
        let fm = FileManager.default
        for d in dirs {
            try? fm.createDirectory(at: d, withIntermediateDirectories: true)
            // Best-effort: make queue folders writable.
            try? fm.setAttributes([.posixPermissions: 0o777], ofItemAtPath: d.path)
        }
    }

    func restartWatchingIncomingFolder() {
        log("Restarting Incoming folder watcher…")

        // Cancel existing sources/timers.
        if let src = dirSource {
            src.cancel()
            dirSource = nil
        }
        if let t = pollTimer {
            t.cancel()
            pollTimer = nil
        }

        stopSecurityScopedAccessIfNeeded()
        resolveAndStartSecurityScopedAccessIfNeeded()

        startWatchingIncomingFolder()
    }

    private func startWatchingIncomingFolder() {
        ensureFolders()
        requestScan(reason: "startup")

        // Safety net: poll periodically. This also keeps things working when filesystem
        // event sources are unavailable (e.g. sandbox restrictions).
        OneNoteHelperWatcherQueue.async { [weak self] in
            guard let self else { return }
            let timer = DispatchSource.makeTimerSource(queue: OneNoteHelperWatcherQueue)
            timer.schedule(deadline: .now() + 2, repeating: .seconds(5))
            timer.setEventHandler { [weak self] in
                self?.requestScan(reason: "poll")
            }
            timer.resume()
            // Keep the timer alive.
            Task { @MainActor in
                self.pollTimer = timer
            }
            // (poller started)
        }

        let dir = folderURL("Incoming")
        let fd = open(dir.path, O_EVTONLY)
        if fd < 0 {
            log("Failed to open Incoming folder: \(dir.path)")
            return
        }

        let source = DispatchSource.makeFileSystemObjectSource(fileDescriptor: fd,
                                                              eventMask: [.write, .rename, .delete, .attrib],
                                                              queue: OneNoteHelperWatcherQueue)
        source.setEventHandler { [weak self] in
            self?.requestScan(reason: "fs-event")
        }
        source.setCancelHandler {
            close(fd)
        }
        self.dirSource = source
        source.resume()
        // (watcher active)
    }

    private func requestScan(reason: String) {
        // Coalesce bursty filesystem events and avoid concurrent scans.
        OneNoteHelperWatcherQueue.async { [weak self] in
            guard let self else { return }

            self.scanWorkItem?.cancel()
            let work = DispatchWorkItem { [weak self] in
                self?.scanIncomingFolder(reason: reason)
            }
            self.scanWorkItem = work
            // Debounce a little: avoids event storms, Finder metadata churn, and our own file operations.
            OneNoteHelperWatcherQueue.asyncAfter(deadline: .now() + 0.35, execute: work)
        }
    }

    private func scanIncomingFolder(reason: String) {
        if isScanning {
            return
        }
        isScanning = true
        defer { isScanning = false }

        autoreleasepool {
            ensureFolders()
            let incoming = folderURL("Incoming")
            let processing = folderURL("Processing")
            // (scan)
            let fm = FileManager.default

            guard let items = try? fm.contentsOfDirectory(at: incoming, includingPropertiesForKeys: nil) else {
                self.log("Incoming scan: failed to list directory \(incoming.path)")
                return
            }

            // Look for pairs: job-*.(pdf|ps) + job-*.json
            let docs = items.filter {
                let ext = $0.pathExtension.lowercased()
                return ext == "pdf" || ext == "ps"
            }
            // (scan results)
            for doc in docs {
                let json = doc.deletingPathExtension().appendingPathExtension("json")
                guard fm.fileExists(atPath: json.path) else {
                    // skipping doc without JSON
                    continue
                }

                let procDoc = processing.appendingPathComponent(doc.lastPathComponent)
                let procJson = processing.appendingPathComponent(json.lastPathComponent)

                do {
                    try self.stageFile(from: doc, to: procDoc)
                    try self.stageFile(from: json, to: procJson)
                    self.log("Queued job: staged \(doc.lastPathComponent) and \(json.lastPathComponent) to Processing")
                } catch {
                    self.log("Incoming scan: stage failed for \(doc.lastPathComponent): \(error.localizedDescription)")
                    continue
                }

                processQueuedJob(pdfURL: procDoc, jsonURL: procJson)
            }
        }
    }

    private struct QueueMeta: Decodable {
        let file: String
        let title: String
        let user: String
        let job: String
    }

    nonisolated private func stageFile(from: URL, to: URL) throws {
        let fm = FileManager.default

        // Ensure destination directory exists.
        try? fm.createDirectory(at: to.deletingLastPathComponent(), withIntermediateDirectories: true)

        // If destination exists, remove it first.
        if fm.fileExists(atPath: to.path) {
            try fm.removeItem(at: to)
        }

        // Copy first.
        try fm.copyItem(at: from, to: to)

        // Then try to remove original (best effort).
        do {
            try fm.removeItem(at: from)
        } catch {
            // We may not be allowed to delete the original if the Incoming directory has the sticky bit
            // and the file is owned by another user. In that case, leave it; we'll rely on de-dupe
            // (destination-exists check) to avoid reprocessing.
            self.log("Stage: warning: copied but could not remove original \(from.lastPathComponent): \(error.localizedDescription)")
        }
    }

    nonisolated private func processQueuedJob(pdfURL: URL, jsonURL: URL) {
        self.log("Processing queued job: pdf=\(pdfURL.lastPathComponent), json=\(jsonURL.lastPathComponent)")
        OneNoteHelperWatcherQueue.async {
            let fm = FileManager.default
            let done = self.folderURL("Done")
            let failed = self.folderURL("Failed")

            guard let data = try? Data(contentsOf: jsonURL),
                  let meta = try? JSONDecoder().decode(QueueMeta.self, from: data) else {
                self.log("Failed to parse metadata: \(jsonURL.path)")
                try? fm.moveItem(at: pdfURL, to: failed.appendingPathComponent(pdfURL.lastPathComponent))
                try? fm.moveItem(at: jsonURL, to: failed.appendingPathComponent(jsonURL.lastPathComponent))
                return
            }
            // (meta)

            Task { @MainActor in
                // (import start)
                self.importFile(filePath: pdfURL.path, title: meta.title, user: meta.user, job: meta.job) { ok in
                    // (import completed)
                    OneNoteHelperWatcherQueue.async {
                        let fm = FileManager.default
                        let target = ok ? done : failed
                        try? fm.moveItem(at: pdfURL, to: target.appendingPathComponent(pdfURL.lastPathComponent))
                        try? fm.moveItem(at: jsonURL, to: target.appendingPathComponent(jsonURL.lastPathComponent))
                    }
                }
            }
        }
    }

    // MARK: - Graph API helpers

    func fetchNotebooks(completion: @escaping (Result<[OneNoteTargetStore.Notebook], Error>) -> Void) {
        acquireGraphToken { token in
            guard let token else {
                completion(.failure(NSError(domain: "OneNoteHelper", code: 401)))
                return
            }

            struct Response: Decodable {
                struct Notebook: Decodable { let id: String; let displayName: String? }
                let value: [Notebook]
                let nextLink: String?
                private enum CodingKeys: String, CodingKey {
                    case value
                    case nextLink = "@odata.nextLink"
                }
            }

            func fetchPage(urlString: String, accum: [OneNoteTargetStore.Notebook]) {
                self.graphGET(token: token, url: urlString) { result in
                    switch result {
                    case .failure(let err):
                        completion(.failure(err))
                    case .success(let data):
                        guard let resp = try? JSONDecoder().decode(Response.self, from: data) else {
                            let payload = String(data: data, encoding: .utf8) ?? ""
                            self.log("Notebooks decode failed. Payload: \(payload)")
                            completion(.failure(NSError(domain: "OneNoteHelper", code: 500)))
                            return
                        }
                        let mapped = resp.value.map { OneNoteTargetStore.Notebook(id: $0.id, name: $0.displayName ?? $0.id) }
                        let total = accum + mapped
                        if let next = resp.nextLink {
                            fetchPage(urlString: next, accum: total)
                        } else {
                            completion(.success(total.sorted { $0.name.localizedCaseInsensitiveCompare($1.name) == .orderedAscending }))
                        }
                    }
                }
            }

            let first = self.graphURL("me/onenote/notebooks?$select=id,displayName&$top=200")
            fetchPage(urlString: first, accum: [])
        }
    }

    func fetchSections(notebookId: String, completion: @escaping (Result<[OneNoteTargetStore.Section], Error>) -> Void) {
        acquireGraphToken { token in
            guard let token else {
                completion(.failure(NSError(domain: "OneNoteHelper", code: 401)))
                return
            }

            struct Response: Decodable {
                struct Section: Decodable { let id: String; let displayName: String? }
                let value: [Section]
                let nextLink: String?
                private enum CodingKeys: String, CodingKey {
                    case value
                    case nextLink = "@odata.nextLink"
                }
            }

            func fetchPage(urlString: String, accum: [OneNoteTargetStore.Section]) {
                self.graphGET(token: token, url: urlString) { result in
                    switch result {
                    case .failure(let err):
                        completion(.failure(err))
                    case .success(let data):
                        guard let resp = try? JSONDecoder().decode(Response.self, from: data) else {
                            let payload = String(data: data, encoding: .utf8) ?? ""
                            self.log("Sections decode failed. Payload: \(payload)")
                            completion(.failure(NSError(domain: "OneNoteHelper", code: 500)))
                            return
                        }
                        let mapped = resp.value.map { OneNoteTargetStore.Section(id: $0.id, name: $0.displayName ?? $0.id) }
                        let total = accum + mapped
                        if let next = resp.nextLink {
                            fetchPage(urlString: next, accum: total)
                        } else {
                            completion(.success(total.sorted { $0.name.localizedCaseInsensitiveCompare($1.name) == .orderedAscending }))
                        }
                    }
                }
            }

            // Includes sections inside section groups.
            let first = self.graphURL("me/onenote/notebooks/\(notebookId)/sections?$select=id,displayName&$top=200")
            fetchPage(urlString: first, accum: [])
        }
    }

    func fetchPages(sectionId: String, completion: @escaping (Result<[OneNoteTargetStore.Page], Error>) -> Void) {
        acquireGraphToken { token in
            guard let token else {
                completion(.failure(NSError(domain: "OneNoteHelper", code: 401)))
                return
            }

            struct Response: Decodable {
                struct Page: Decodable { let id: String; let title: String? }
                let value: [Page]
                let nextLink: String?
                private enum CodingKeys: String, CodingKey {
                    case value
                    case nextLink = "@odata.nextLink"
                }
            }

            func fetchPage(urlString: String, accum: [OneNoteTargetStore.Page]) {
                self.graphGET(token: token, url: urlString) { result in
                    switch result {
                    case .failure(let err):
                        completion(.failure(err))
                    case .success(let data):
                        guard let resp = try? JSONDecoder().decode(Response.self, from: data) else {
                            let payload = String(data: data, encoding: .utf8) ?? ""
                            self.log("Pages decode failed. Payload: \(payload)")
                            completion(.failure(NSError(domain: "OneNoteHelper", code: 500)))
                            return
                        }
                        let mapped = resp.value.map { OneNoteTargetStore.Page(id: $0.id, title: $0.title ?? $0.id) }
                        let total = accum + mapped
                        if let next = resp.nextLink {
                            fetchPage(urlString: next, accum: total)
                        } else {
                            completion(.success(total.sorted { $0.title.localizedCaseInsensitiveCompare($1.title) == .orderedAscending }))
                        }
                    }
                }
            }

            let first = self.graphURL("me/onenote/sections/\(sectionId)/pages?$select=id,title&$top=100")
            fetchPage(urlString: first, accum: [])
        }
    }

    nonisolated private func graphGET(token: String, url: String, completion: @escaping (Result<Data, Error>) -> Void) {
        guard let u = URL(string: url) else {
            completion(.failure(NSError(domain: "OneNoteHelper", code: 400)))
            return
        }
        var request = URLRequest(url: u)
        request.httpMethod = "GET"
        request.setValue("Bearer \(token)", forHTTPHeaderField: "Authorization")

        URLSession.shared.dataTask(with: request) { data, response, error in
            if let error = error {
                self.log("Graph GET failed: \(error.localizedDescription)")
                completion(.failure(error))
                return
            }
            if let http = response as? HTTPURLResponse, http.statusCode >= 300 {
                let payload = data.flatMap { String(data: $0, encoding: .utf8) } ?? ""
                self.log("Graph GET failed (\(http.statusCode)): \(payload)")
                completion(.failure(NSError(domain: "OneNoteHelper", code: http.statusCode)))
                return
            }
            completion(.success(data ?? Data()))
        }.resume()
    }

    func signIn(completion: ((Bool) -> Void)? = nil) {
        guard let application = msalApplication else {
            log("MSAL is not configured. Please enter MSAL Client ID and Redirect URI.")
            completion?(false)
            return
        }

        log("Starting interactive sign-in…")
        acquireGraphTokenInteractive(application: application) { token in
            if let token {
                self.log("Sign-in OK (token len=\(token.count))")
                completion?(true)
            } else {
                self.log("Sign-in FAILED")
                completion?(false)
            }
        }
    }

    private func acquireGraphToken(completion: @escaping (String?) -> Void) {
        guard let application = msalApplication else {
            let bundleId = Bundle.main.bundleIdentifier ?? "(nil)"
            log("MSAL is not configured. bundleId=\(bundleId)")
            log("MSALClientId (settings) = \(self.msalClientId ?? "")")
            log("MSALRedirectUri (settings) = \(self.msalRedirectUri ?? "")")
            log("TIP: Add a URL type in Info.plist with scheme 'msauth.\(Bundle.main.bundleIdentifier ?? "")' so MSAL can handle the redirect.")
            completion(nil)
            return
        }

        let effAuthority = (self.msalAuthority?.trimmingCharacters(in: .whitespacesAndNewlines)).flatMap { $0.isEmpty ? nil : $0 } ?? "https://login.microsoftonline.com/organizations"
        let effRedirect = "msauth.\(Bundle.main.bundleIdentifier ?? "")://auth"
        self.log("MSAL effective settings: authority=\(effAuthority), redirect=\(effRedirect)")
        
        do {
            let accounts = try application.allAccounts()
            self.log("MSAL accounts count: \(accounts.count)")
            if let account = accounts.first {
                let parameters = MSALSilentTokenParameters(scopes: graphScopes, account: account)
                application.acquireTokenSilent(with: parameters) { result, error in
                    self.log("Attempting silent token acquisition…")
                    if let result = result {
                        completion(result.accessToken)
                        return
                    }
                    if let error = error {
                        let nserr = error as NSError
                        self.log("MSAL silent token error: domain=\(nserr.domain) code=\(nserr.code) desc=\(nserr.localizedDescription)")
                    }
                    Task { @MainActor in
                        self.acquireGraphTokenInteractive(application: application, completion: completion)
                    }
                }
                return
            }
        } catch {
            log("MSAL failed to load accounts: \(error.localizedDescription)")
        }

        acquireGraphTokenInteractive(application: application, completion: completion)
    }

    private func acquireGraphTokenInteractive(application: MSALPublicClientApplication, completion: @escaping (String?) -> Void) {
        // Coalesce concurrent interactive auth requests; MSAL allows only one at a time.
        if interactiveAuthInProgress {
            log("MSAL interactive auth already in progress; queuing request…")
            interactiveAuthWaiters.append(completion)
            return
        }
        interactiveAuthInProgress = true
        interactiveAuthWaiters.append(completion)

        // Ensure we have a normal UI so MSAL can present.
        if NSApp.activationPolicy() != .regular {
            _ = NSApp.setActivationPolicy(.regular)
        }
        NSApp.activate(ignoringOtherApps: true)

        guard let presentationVC = NSApp.keyWindow?.contentViewController ?? NSApp.windows.first?.contentViewController else {
            log("MSAL interactive auth requires a window")
            let waiters = self.interactiveAuthWaiters
            self.interactiveAuthWaiters.removeAll()
            self.interactiveAuthInProgress = false
            for waiter in waiters {
                waiter(nil)
            }
            return
        }

        let authString = (self.msalAuthority?.trimmingCharacters(in: .whitespacesAndNewlines)).flatMap { $0.isEmpty ? nil : $0 } ?? "https://login.microsoftonline.com/organizations"
        self.log("Starting interactive acquisition with authority=\(authString)")
        let webParameters = MSALWebviewParameters(authPresentationViewController: presentationVC)
        let parameters = MSALInteractiveTokenParameters(scopes: graphScopes, webviewParameters: webParameters)
        if let url = URL(string: authString), let auth = try? MSALAADAuthority(url: url) {
            parameters.authority = auth
        }
        parameters.promptType = .selectAccount
        application.acquireToken(with: parameters) { result, error in
            Task { @MainActor in
                let token = result?.accessToken
                if let error = error {
                    let nserr = error as NSError
                    self.log("MSAL interactive token error: domain=\(nserr.domain) code=\(nserr.code) desc=\(nserr.localizedDescription)")
                }
                let waiters = self.interactiveAuthWaiters
                self.interactiveAuthWaiters.removeAll()
                self.interactiveAuthInProgress = false
                for waiter in waiters {
                    waiter(token)
                }
            }
        }
    }

    nonisolated private func uploadSinglePage(token: String, filePath: String, title: String, user: String, job: String, completion: @escaping (Bool) -> Void) {
        // Target page title for all print jobs.
        let pageTitle = "Sent To OneNote"
        let jobTitle = title.isEmpty ? "Printed Document" : title
        let fileURL = URL(fileURLWithPath: filePath)

        // Print system may provide PostScript content even if the file is named .pdf.
        // Detect and convert to real PDF first.
        var psConverted = false
        let effectiveURL: URL
        if self.isPostScript(fileURL: fileURL) {
            if let converted = self.convertPostScriptToPDF(fileURL: fileURL) {
                self.log("Converted PostScript to PDF: \(converted.lastPathComponent) (from \(fileURL.lastPathComponent))")
                effectiveURL = converted
                psConverted = true
            } else {
                self.log("ERROR: PostScript->PDF conversion failed for \(fileURL.path)")
                completion(false)
                return
            }
        } else {
            effectiveURL = fileURL
        }

        enum ImportMode: String {
            case image
            case text
            case hybrid
        }

        let importModeRaw = (UserDefaults.standard.string(forKey: "ImportMode") ?? "hybrid").trimmingCharacters(in: .whitespacesAndNewlines).lowercased()
        let importMode = ImportMode(rawValue: importModeRaw) ?? .hybrid

        // Prefer selectable text: extract attributed text from the PDF and convert to HTML.
        // Depending on ImportMode, we may force images-only or text-only.
        // If extraction fails (e.g. scanned PDF), Hybrid can fall back to rendering pages as images.
        let maxPages = 200
        let renderScale: CGFloat = 2.0

        let boundary = "----onenote-\(UUID().uuidString)"
        var body = Data()

        let targetSectionId = UserDefaults.standard.string(forKey: targetSectionIdKey)
        let targetPageId = UserDefaults.standard.string(forKey: targetPageIdKey)
        let shouldAppendToPage = (targetPageId?.isEmpty == false)

        struct OneNotePatchCommand: Encodable {
            let target: String
            let action: String
            let content: String
        }

        func appendMainPartForCreate(htmlDocument: String) {
            appendMultipartPart(&body, boundary: boundary,
                                contentType: "text/html; charset=utf-8",
                                contentDisposition: "form-data; name=\"Presentation\"",
                                data: Data(htmlDocument.utf8))
        }

        func appendMainPartForAppend(htmlFragment: String) {
            // Use OneNote patch commands to append HTML fragment to the end of the page.
            let commands = [OneNotePatchCommand(target: "body", action: "append", content: htmlFragment)]
            let data = (try? JSONEncoder().encode(commands)) ?? Data("[]".utf8)
            appendMultipartPart(&body, boundary: boundary,
                                contentType: "application/json; charset=utf-8",
                                contentDisposition: "form-data; name=\"commands\"",
                                data: data)
        }

        func fallbackToImages() -> Bool {
            guard let images = self.renderPDFAsPNGs(fileURL: effectiveURL, maxPages: min(maxPages, 30), scale: renderScale), !images.isEmpty else {
                self.log("Failed to extract text or render PDF at \(filePath)")
                return false
            }

            self.log("Upload: preparing IMAGE page for sectionId=\(UserDefaults.standard.string(forKey: targetSectionIdKey) ?? "(default)") title=\(pageTitle) renderedPages=\(images.count)")

            var imgHTML = ""
            for (idx, item) in images.enumerated() {
                imgHTML += "<div style=\"margin: 12px 0;\"><img src=\"name:\(item.token)\" alt=\"Page \(idx + 1)\" /></div>\n"
            }

            let html = """
            <!DOCTYPE html>
            <html>
              <head>
                <meta charset="utf-8" />
                <title>\(escapeHTML(jobTitle))</title>
              </head>
              <body>
                <h1>\(escapeHTML(jobTitle))</h1>
                <p><b>Source:</b> \(escapeHTML(pageTitle))</p>
                <p>Imported by OneNote Helper.</p>
                <p>User: \(escapeHTML(user)) &nbsp; Job: \(escapeHTML(job))</p>
                \(imgHTML)
              </body>
            </html>
            """

            if shouldAppendToPage {
                let fragment = """
                <div>
                  <h2>\(escapeHTML(jobTitle))</h2>
                  <p><b>Source:</b> \(escapeHTML(pageTitle))</p>
                  <p>Imported by OneNote Helper.</p>
                  <p>User: \(escapeHTML(user)) &nbsp; Job: \(escapeHTML(job))</p>
                  \(imgHTML)
                </div>
                """
                appendMainPartForAppend(htmlFragment: fragment)
            } else {
                appendMainPartForCreate(htmlDocument: html)
            }

            for item in images {
                appendMultipartPart(&body,
                                    boundary: boundary,
                                    contentType: "image/png",
                                    contentDisposition: "form-data; name=\"\(item.token)\"; filename=\"\(item.filename)\"",
                                    data: item.data)
            }

            return true
        }

        switch importMode {
        case .image:
            self.log("Import mode=Image; forcing rendered pages as PNGs")
            if !fallbackToImages() {
                completion(false)
                return
            }

        case .text:
            self.log("Import mode=Text; extracting text only (no images)")
            guard let pagesHTMLRaw = self.extractPDFPagesAsHTMLBodies(fileURL: effectiveURL, maxPages: maxPages) else {
                self.log("ERROR: No extracted HTML (text mode) for \(filePath)")
                completion(false)
                return
            }

            let joinedHTML = pagesHTMLRaw.joined(separator: "\n<hr />\n")
            let extractedHTML = joinedHTML.trimmingCharacters(in: .whitespacesAndNewlines)
            self.log("Extracted HTML bytes=\(extractedHTML.utf8.count)")

            guard !extractedHTML.isEmpty else {
                self.log("ERROR: Extracted HTML empty (text mode) for \(filePath)")
                completion(false)
                return
            }

            let preview = String(extractedHTML.prefix(200)).replacingOccurrences(of: "\n", with: "\\n")
            self.log("Extracted HTML preview: \(preview)")

            // Text-only mode: upload just the extracted text. Do not embed images and do not fall back to rendered pages.
            var pageSections = ""
            for (idx, pageHTMLBody) in pagesHTMLRaw.enumerated() {
                pageSections += "<h2>Page \(idx + 1)</h2>\n"
                pageSections += pageHTMLBody
                if idx < pagesHTMLRaw.count - 1 {
                    pageSections += "\n<hr />\n"
                }
            }

            let html = """
            <!DOCTYPE html>
            <html>
              <head>
                <meta charset="utf-8" />
                <title>\(escapeHTML(jobTitle))</title>
              </head>
              <body>
                <h1>\(escapeHTML(jobTitle))</h1>
                <p><b>Source:</b> \(escapeHTML(pageTitle))</p>
                <p>Imported by OneNote Helper.</p>
                <p>User: \(escapeHTML(user)) &nbsp; Job: \(escapeHTML(job))</p>
                <hr />
                \(pageSections)
              </body>
            </html>
            """

            if shouldAppendToPage {
                let fragment = """
                <div>
                  <h2>\(escapeHTML(jobTitle))</h2>
                  <p><b>Source:</b> \(escapeHTML(pageTitle))</p>
                  <p>Imported by OneNote Helper.</p>
                  <p>User: \(escapeHTML(user)) &nbsp; Job: \(escapeHTML(job))</p>
                  <hr />
                  \(pageSections)
                </div>
                """
                appendMainPartForAppend(htmlFragment: fragment)
            } else {
                appendMainPartForCreate(htmlDocument: html)
            }

        case .hybrid:
            if let pagesHTMLRaw = self.extractPDFPagesAsHTMLBodies(fileURL: effectiveURL, maxPages: maxPages) {
                // Join to run heuristics + logs.
                let joinedHTML = pagesHTMLRaw.joined(separator: "\n<hr />\n")
                let extractedHTML = joinedHTML.trimmingCharacters(in: .whitespacesAndNewlines)
                self.log("Extracted HTML bytes=\(extractedHTML.utf8.count)")
                if !extractedHTML.isEmpty {
                    let preview = String(extractedHTML.prefix(200)).replacingOccurrences(of: "\n", with: "\\n")
                    self.log("Extracted HTML preview: \(preview)")

                    // Heuristic: PDFs converted from PostScript often have garbage text extraction (missing ToUnicode).
                    // In hybrid mode, we keep images, but replace the extracted text with a short note if it is gibberish.
                    var effectivePagesHTML = pagesHTMLRaw
                    if psConverted {
                        let plain = self.plainTextFromHTML(extractedHTML)
                        if self.textLooksGibberish(plain) {
                            self.log("Extracted text looks like gibberish (PS-converted); replacing text with note (hybrid mode)")
                            effectivePagesHTML = Array(repeating: "<p><i>(Text extraction from this print job is unreliable; see images below.)</i></p>", count: pagesHTMLRaw.count)
                        }
                    }

                    self.log("Upload: preparing HYBRID (per-page) page for sectionId=\(UserDefaults.standard.string(forKey: targetSectionIdKey) ?? "(default)") title=\(pageTitle)")

                    // Hybrid mode: include extracted text + embedded PDF image XObjects, placed after the page text.
                    let xobjImages = self.extractPDFImageXObjects(fileURL: effectiveURL, maxPages: min(maxPages, 200), maxImages: 400)
                    var imagesByPage: [Int: [EmbeddedImagePart]] = [:]
                    for img in xobjImages {
                        imagesByPage[img.pageIndex, default: []].append(img)
                    }

                    if !xobjImages.isEmpty {
                        self.log("Found \(xobjImages.count) PDF image XObject(s); embedding as attachments")
                    }

                    var pageSections = ""
                    for (idx, pageHTMLBody) in effectivePagesHTML.enumerated() {
                        let imgs = imagesByPage[idx] ?? []
                        var imgsHTML = ""
                        if !imgs.isEmpty {
                            imgsHTML += "\n<div style=\"margin-top: 12px;\">\n"
                            for (j, item) in imgs.enumerated() {
                                imgsHTML += "<div style=\"margin: 10px 0;\"><img src=\"name:\(item.token)\" alt=\"Image \(j + 1)\" /></div>\n"
                            }
                            imgsHTML += "</div>\n"
                        }

                        pageSections += "<h2>Page \(idx + 1)</h2>\n"
                        pageSections += pageHTMLBody
                        pageSections += imgsHTML
                        if idx < pagesHTMLRaw.count - 1 {
                            pageSections += "\n<hr />\n"
                        }
                    }

                    let html = """
                    <!DOCTYPE html>
                    <html>
                      <head>
                        <meta charset="utf-8" />
                        <title>\(escapeHTML(jobTitle))</title>
                      </head>
                      <body>
                        <h1>\(escapeHTML(jobTitle))</h1>
                        <p><b>Source:</b> \(escapeHTML(pageTitle))</p>
                        <p>Imported by OneNote Helper.</p>
                        <p>User: \(escapeHTML(user)) &nbsp; Job: \(escapeHTML(job))</p>
                        <hr />
                        \(pageSections)
                      </body>
                    </html>
                    """

                    if shouldAppendToPage {
                        let fragment = """
                        <div>
                          <h2>\(escapeHTML(jobTitle))</h2>
                          <p><b>Source:</b> \(escapeHTML(pageTitle))</p>
                          <p>Imported by OneNote Helper.</p>
                          <p>User: \(escapeHTML(user)) &nbsp; Job: \(escapeHTML(job))</p>
                          <hr />
                          \(pageSections)
                        </div>
                        """
                        appendMainPartForAppend(htmlFragment: fragment)
                    } else {
                        appendMainPartForCreate(htmlDocument: html)
                    }

                    // Attach embedded images referenced by name:<token>.
                    for item in xobjImages {
                        appendMultipartPart(&body,
                                            boundary: boundary,
                                            contentType: item.mimeType,
                                            contentDisposition: "form-data; name=\"\(item.token)\"; filename=\"\(item.filename)\"",
                                            data: item.data)
                    }
                } else {
                    self.log("Extracted HTML empty; falling back to images")
                    if !fallbackToImages() {
                        completion(false)
                        return
                    }
                    // Continue to Graph upload below.
                }
            } else {
                self.log("No extracted HTML; falling back to images")
                if !fallbackToImages() {
                    completion(false)
                    return
                }
                // Continue to Graph upload below.
            }
        }

        body.append("--\(boundary)--\r\n".data(using: .utf8)!)

        let urlString: String
        let method: String
        if shouldAppendToPage, let targetPageId, !targetPageId.isEmpty {
            urlString = self.graphURL("me/onenote/pages/\(targetPageId)/content")
            method = "PATCH"
        } else if let targetSectionId, !targetSectionId.isEmpty {
            urlString = self.graphURL("me/onenote/sections/\(targetSectionId)/pages")
            method = "POST"
        } else {
            // Fallback to default notebook/section behavior.
            urlString = self.graphURL("me/onenote/pages")
            method = "POST"
        }

        guard let url = URL(string: urlString) else {
            log("Invalid Graph URL")
            completion(false)
            return
        }

        var request = URLRequest(url: url)
        request.httpMethod = method
        request.setValue("Bearer \(token)", forHTTPHeaderField: "Authorization")
        request.setValue("multipart/form-data; boundary=\(boundary)", forHTTPHeaderField: "Content-Type")
        request.setValue("application/json", forHTTPHeaderField: "Accept")
        request.httpBody = body

        let task = URLSession.shared.dataTask(with: request) { data, response, error in
            if let error = error {
                self.log("Graph upload failed: \(error.localizedDescription)")
                completion(false)
                return
            }

            if let http = response as? HTTPURLResponse, http.statusCode >= 300 {
                let payload = data.flatMap { String(data: $0, encoding: .utf8) } ?? ""
                self.log("Graph upload failed (\(http.statusCode)): \(payload)")
                completion(false)
                return
            }

            self.log("Graph upload succeeded")
            completion(true)
        }
        task.resume()
    }

    private struct RenderedPart {
        let token: String
        let filename: String
        let data: Data
    }

    private struct EmbeddedImagePart {
        let pageIndex: Int
        let token: String
        let filename: String
        let mimeType: String
        let data: Data
    }

    nonisolated private func isPostScript(fileURL: URL) -> Bool {
        guard let fh = try? FileHandle(forReadingFrom: fileURL) else { return false }
        defer { try? fh.close() }
        let head = (try? fh.read(upToCount: 16)) ?? Data()
        if let s = String(data: head, encoding: .ascii) {
            return s.hasPrefix("%!PS")
        }
        return false
    }

    nonisolated private func convertPostScriptToPDF(fileURL: URL) -> URL? {
        // Only supported path: delegate PS->PDF conversion to the embedded XPC service (Ghostscript).
        // CoreGraphics PS conversion was unreliable in practice, and executing system/Homebrew tools
        // is not compatible with the App Sandbox.
        return convertPostScriptToPDFUsingBundledGhostscript(fileURL: fileURL)
    }

    nonisolated private func convertPostScriptToPDFUsingBundledGhostscript(fileURL: URL) -> URL? {
        // In the sandboxed app, we cannot exec gs directly (it gets SIGKILL).
        // We delegate to an embedded XPC service which runs gs out-of-sandbox.

        let tmpURL = URL(fileURLWithPath: NSTemporaryDirectory(), isDirectory: true)
            .appendingPathComponent("onenotehelper-\(UUID().uuidString).pdf")

        self.log("PS->PDF: requesting XPC conversion")
        let res = GhostscriptXPCClient.convertPS(psPath: fileURL.path, pdfPath: tmpURL.path, timeoutSeconds: 90)

        if !res.logs.isEmpty {
            self.log("PS->PDF(XPC) logs: \(res.logs)")
        }

        guard res.ok else {
            self.log("PS->PDF(XPC) failed")
            return nil
        }

        let outData = (try? Data(contentsOf: tmpURL)) ?? Data()
        if !outData.isEmpty, PDFDocument(data: outData) != nil {
            self.log("PS->PDF conversion used XPC gs")
            return tmpURL
        }

        self.log("PS->PDF(XPC): produced invalid/empty PDF")
        return nil
    }


    nonisolated private func plainTextFromHTML(_ html: String) -> String {
        // Very small HTML stripper for heuristic purposes.
        // This doesn't need to be perfect; we just want to detect obvious gibberish.
        var s = html
        // Remove tags
        s = s.replacingOccurrences(of: "<[^>]+>", with: " ", options: .regularExpression)
        // Decode a couple of common entities
        s = s.replacingOccurrences(of: "&nbsp;", with: " ")
        s = s.replacingOccurrences(of: "&amp;", with: "&")
        s = s.replacingOccurrences(of: "&lt;", with: "<")
        s = s.replacingOccurrences(of: "&gt;", with: ">")
        s = s.replacingOccurrences(of: "&quot;", with: "\"")
        // Collapse whitespace
        s = s.replacingOccurrences(of: "\\s+", with: " ", options: .regularExpression)
        return s.trimmingCharacters(in: .whitespacesAndNewlines)
    }

    nonisolated private func textLooksGibberish(_ text: String) -> Bool {
        let t = text.trimmingCharacters(in: .whitespacesAndNewlines)
        if t.isEmpty { return true }

        // Look at a prefix to keep it cheap.
        let sample = String(t.prefix(2000))
        let total = sample.count
        if total < 20 { return true }

        let lettersDigits = sample.unicodeScalars.filter { CharacterSet.alphanumerics.contains($0) }.count
        let spaces = sample.unicodeScalars.filter { CharacterSet.whitespacesAndNewlines.contains($0) }.count
        let punct = total - lettersDigits - spaces

        let alnumRatio = Double(lettersDigits) / Double(total)
        let punctRatio = Double(punct) / Double(total)

        // Heuristics tuned for the observed output: lots of punctuation, few letters.
        if alnumRatio < 0.35 { return true }
        if punctRatio > 0.40 { return true }

        // Also flag if there are almost no letters (only digits/symbols)
        let letters = sample.unicodeScalars.filter { CharacterSet.letters.contains($0) }.count
        if Double(letters) / Double(total) < 0.15 { return true }

        return false
    }

    nonisolated private func renderPDFAsPNGs(fileURL: URL, maxPages: Int, scale: CGFloat) -> [RenderedPart]? {
        // Prefer dataRepresentation load to avoid file coordination/sandbox oddities.
        let doc: PDFDocument?
        if let data = try? Data(contentsOf: fileURL) {
            doc = PDFDocument(data: data)
        } else {
            doc = PDFDocument(url: fileURL)
        }
        guard let doc else { return nil }
        let pageCount = min(doc.pageCount, maxPages)
        if pageCount <= 0 { return [] }

        var parts: [RenderedPart] = []
        parts.reserveCapacity(pageCount)

        for i in 0..<pageCount {
            guard let page = doc.page(at: i) else { continue }
            let bounds = page.bounds(for: .mediaBox)
            let targetSize = CGSize(width: max(1, bounds.width * scale), height: max(1, bounds.height * scale))

            let image = NSImage(size: targetSize)
            image.lockFocusFlipped(false)
            if let ctx = NSGraphicsContext.current?.cgContext {
                ctx.setFillColor(NSColor.white.cgColor)
                ctx.fill(CGRect(origin: .zero, size: targetSize))
                ctx.saveGState()
                ctx.scaleBy(x: scale, y: scale)
                page.draw(with: .mediaBox, to: ctx)
                ctx.restoreGState()
            }
            image.unlockFocus()

            guard let tiff = image.tiffRepresentation,
                  let rep = NSBitmapImageRep(data: tiff),
                  let png = rep.representation(using: .png, properties: [:]) else {
                continue
            }

            let token = "img\(i + 1)"
            let filename = String(format: "page-%03d.png", i + 1)
            parts.append(RenderedPart(token: token, filename: filename, data: png))
        }

        return parts
    }

    // MARK: - PDF image XObject extraction (hybrid upload)

    nonisolated private func extractPDFImageXObjects(fileURL: URL, maxPages: Int, maxImages: Int) -> [EmbeddedImagePart] {
        guard let doc = CGPDFDocument(fileURL as CFURL) else { return [] }

        let pageCount = min(doc.numberOfPages, maxPages)
        if pageCount <= 0 { return [] }

        final class Box {
            var images: [EmbeddedImagePart] = []
            var pageIndex: Int = 0
        }
        let box = Box()

        func parseXObjectStream(_ stream: CGPDFStreamRef) {
            // Note: we intentionally do not try to de-duplicate streams by pointer identity here.
            // CoreGraphics stream types are not stable Swift pointer types across SDKs.

            guard let dict = CGPDFStreamGetDictionary(stream) else { return }

            var subtypeName: UnsafePointer<Int8>?
            if CGPDFDictionaryGetName(dict, "Subtype", &subtypeName), let subtypeName {
                let subtype = String(cString: subtypeName)
                if subtype == "Image" {
                    // Determine filter to decide output encoding.
                    var filterName: UnsafePointer<Int8>?
                    let filter: String? = {
                        if CGPDFDictionaryGetName(dict, "Filter", &filterName), let filterName {
                            return String(cString: filterName)
                        }
                        return nil
                    }()

                    // Extract stream data.
                    var format: CGPDFDataFormat = .raw
                    guard let cfData = CGPDFStreamCopyData(stream, &format) else { return }
                    let rawData = cfData as Data
                    if rawData.isEmpty { return }

                    func intValue(_ key: String) -> Int? {
                        var v: CGPDFInteger = 0
                        if CGPDFDictionaryGetInteger(dict, key, &v) {
                            return Int(v)
                        }
                        return nil
                    }

                    func colorSpaceName() -> String? {
                        // ColorSpace can be a name or an array (first element is name)
                        var csObj: CGPDFObjectRef?
                        if CGPDFDictionaryGetObject(dict, "ColorSpace", &csObj), let csObj {
                            var namePtr: UnsafePointer<Int8>?
                            if CGPDFObjectGetValue(csObj, .name, &namePtr), let namePtr {
                                return String(cString: namePtr)
                            }
                            var arr: CGPDFArrayRef?
                            if CGPDFObjectGetValue(csObj, .array, &arr), let arr {
                                var firstObj: CGPDFObjectRef?
                                if CGPDFArrayGetObject(arr, 0, &firstObj), let firstObj {
                                    var n2: UnsafePointer<Int8>?
                                    if CGPDFObjectGetValue(firstObj, .name, &n2), let n2 {
                                        return String(cString: n2)
                                    }
                                }
                            }
                        }
                        return nil
                    }

                    let token = String(format: "xobj_p%03d_%03d", box.pageIndex + 1, box.images.count + 1)

                    // 1) JPEG / JP2 images: embed as-is.
                    if filter == "DCTDecode" {
                        let filename = "\(token).jpg"
                        if box.images.count < maxImages {
                            box.images.append(EmbeddedImagePart(pageIndex: box.pageIndex, token: token, filename: filename, mimeType: "image/jpeg", data: rawData))
                        }
                        return
                    }
                    if filter == "JPXDecode" {
                        let filename = "\(token).jp2"
                        if box.images.count < maxImages {
                            box.images.append(EmbeddedImagePart(pageIndex: box.pageIndex, token: token, filename: filename, mimeType: "image/jp2", data: rawData))
                        }
                        return
                    }

                    // 2) FlateDecode (or no filter): raw bitmap data. Try to decode and re-encode as PNG.
                    if filter == "FlateDecode" || filter == nil {
                        guard let w = intValue("Width"), let h = intValue("Height"), w > 0, h > 0 else { return }
                        let bpc = intValue("BitsPerComponent") ?? 8
                        guard bpc == 8 else {
                            // Keep scope small: implement 1-bit/16-bit later if needed.
                            return
                        }

                        let cs = colorSpaceName() ?? "DeviceRGB"
                        let components: Int
                        let cgColorSpace: CGColorSpace
                        if cs == "DeviceGray" {
                            components = 1
                            cgColorSpace = CGColorSpaceCreateDeviceGray()
                        } else if cs == "DeviceRGB" {
                            components = 3
                            cgColorSpace = CGColorSpaceCreateDeviceRGB()
                        } else {
                            // Unsupported color space for now.
                            return
                        }

                        let expectedMin = w * h * components
                        if rawData.count < expectedMin {
                            // Data is smaller than expected: likely needs predictors/extra decoding.
                            return
                        }

                        let bytesPerRow = w * components
                        guard let provider = CGDataProvider(data: rawData as CFData) else { return }
                        let bitmapInfo: CGBitmapInfo = cs == "DeviceGray" ? [] : CGBitmapInfo(rawValue: CGImageAlphaInfo.none.rawValue)

                        if let cgImage = CGImage(width: w,
                                                 height: h,
                                                 bitsPerComponent: 8,
                                                 bitsPerPixel: components * 8,
                                                 bytesPerRow: bytesPerRow,
                                                 space: cgColorSpace,
                                                 bitmapInfo: bitmapInfo,
                                                 provider: provider,
                                                 decode: nil,
                                                 shouldInterpolate: true,
                                                 intent: .defaultIntent) {
                            let outData = NSMutableData()
                            if let dest = CGImageDestinationCreateWithData(outData, UTType.png.identifier as CFString, 1, nil) {
                                CGImageDestinationAddImage(dest, cgImage, nil)
                                if CGImageDestinationFinalize(dest) {
                                    let filename = "\(token).png"
                                    if box.images.count < maxImages {
                                        box.images.append(EmbeddedImagePart(pageIndex: box.pageIndex, token: token, filename: filename, mimeType: "image/png", data: outData as Data))
                                    }
                                    return
                                }
                            }
                        }
                        return
                    }

                    // Other filters are currently not supported (would require more decoding).
                    return
                } else if subtype == "Form" {
                    // Form XObject can contain nested XObjects.
                    // Recurse into its Resources if present.
                    var resObj: CGPDFObjectRef?
                    if CGPDFDictionaryGetObject(dict, "Resources", &resObj), let resObj {
                        var resDict: CGPDFDictionaryRef?
                        if CGPDFObjectGetValue(resObj, .dictionary, &resDict), let resDict {
                            parseResources(resDict)
                        }
                    }
                }
            }
        }

        func applyDict(_ dict: CGPDFDictionaryRef, _ handler: @escaping (String, CGPDFObjectRef) -> Void) {
            typealias Applier = @convention(c) (UnsafePointer<Int8>, CGPDFObjectRef, UnsafeMutableRawPointer?) -> Void
            let cb: Applier = { keyC, obj, info in
                guard let info else { return }
                let handler = Unmanaged<AnyObject>.fromOpaque(info).takeUnretainedValue() as! (String, CGPDFObjectRef) -> Void
                handler(String(cString: keyC), obj)
            }
            let unmanaged = Unmanaged.passRetained(handler as AnyObject)
            CGPDFDictionaryApplyFunction(dict, cb, unmanaged.toOpaque())
            unmanaged.release()
        }

        func parseResources(_ resources: CGPDFDictionaryRef) {
            var xobjObj: CGPDFObjectRef?
            if !CGPDFDictionaryGetObject(resources, "XObject", &xobjObj) { return }
            guard let xobjObj else { return }

            var xobjDict: CGPDFDictionaryRef?
            if !CGPDFObjectGetValue(xobjObj, .dictionary, &xobjDict) { return }
            guard let xobjDict else { return }

            applyDict(xobjDict) { _, obj in
                var stream: CGPDFStreamRef?
                if CGPDFObjectGetValue(obj, .stream, &stream), let stream {
                    parseXObjectStream(stream)
                }
            }
        }

        for p in 1...pageCount {
            guard let page = doc.page(at: p) else { continue }
            box.pageIndex = p - 1
            guard let pageDict = page.dictionary else { continue }

            // Resources may be inherited, but CGPDFPageGetDictionary should reflect merged dictionary.
            var resObj: CGPDFObjectRef?
            if CGPDFDictionaryGetObject(pageDict, "Resources", &resObj), let resObj {
                var resDict: CGPDFDictionaryRef?
                if CGPDFObjectGetValue(resObj, .dictionary, &resDict), let resDict {
                    parseResources(resDict)
                }
            }

            if box.images.count >= maxImages { break }
        }

        // Best-effort: filter out non-image payloads.
        // Keep only image/* mime types.
        return box.images.filter { $0.mimeType.hasPrefix("image/") }
    }

    nonisolated private func extractPDFAsHTMLBody(fileURL: URL, maxPages: Int) -> String? {
        // Prefer dataRepresentation load to avoid file coordination/sandbox oddities.
        if let data = try? Data(contentsOf: fileURL), let doc = PDFDocument(data: data) {
            return extractPDFDocAsHTMLBody(doc: doc, maxPages: maxPages)
        }
        guard let doc = PDFDocument(url: fileURL) else { return nil }
        return extractPDFDocAsHTMLBody(doc: doc, maxPages: maxPages)
    }

    nonisolated private func extractPDFPagesAsHTMLBodies(fileURL: URL, maxPages: Int) -> [String]? {
        // Prefer dataRepresentation load to avoid file coordination/sandbox oddities.
        let doc: PDFDocument?
        if let data = try? Data(contentsOf: fileURL) {
            doc = PDFDocument(data: data)
        } else {
            doc = PDFDocument(url: fileURL)
        }
        guard let doc else { return nil }

        let pageCount = min(doc.pageCount, maxPages)
        if pageCount <= 0 { return [] }

        var out: [String] = []
        out.reserveCapacity(pageCount)

        for i in 0..<pageCount {
            guard let page = doc.page(at: i) else {
                out.append("")
                continue
            }
            if let a = page.attributedString, a.length > 0 {
                let opts: [NSAttributedString.DocumentAttributeKey: Any] = [
                    .documentType: NSAttributedString.DocumentType.html,
                    .characterEncoding: String.Encoding.utf8.rawValue
                ]
                if let data = try? a.data(from: NSRange(location: 0, length: a.length), documentAttributes: opts),
                   let html = String(data: data, encoding: .utf8) {
                    // Extract <body>...</body> if present.
                    if let bodyStart = html.range(of: "<body"),
                       let bodyTagEnd = html.range(of: ">", range: bodyStart.upperBound..<html.endIndex),
                       let bodyEnd = html.range(of: "</body>") {
                        out.append(String(html[bodyTagEnd.upperBound..<bodyEnd.lowerBound]))
                    } else {
                        out.append(html)
                    }
                } else {
                    out.append("")
                }
            } else {
                out.append("")
            }
        }

        return out
    }

    nonisolated private func extractPDFDocAsHTMLBody(doc: PDFDocument, maxPages: Int) -> String? {
        let pageCount = min(doc.pageCount, maxPages)
        if pageCount <= 0 { return "" }

        let combined = NSMutableAttributedString()

        for i in 0..<pageCount {
            guard let page = doc.page(at: i) else { continue }
            if let a = page.attributedString {
                combined.append(a)
            }
            if i < pageCount - 1 {
                combined.append(NSAttributedString(string: "\n\n"))
            }
        }

        if combined.length == 0 { return nil }

        let opts: [NSAttributedString.DocumentAttributeKey: Any] = [
            .documentType: NSAttributedString.DocumentType.html,
            .characterEncoding: String.Encoding.utf8.rawValue
        ]

        guard let data = try? combined.data(from: NSRange(location: 0, length: combined.length), documentAttributes: opts),
              let html = String(data: data, encoding: .utf8) else {
            return nil
        }

        // Extract <body>...</body> if present; otherwise return whole HTML.
        if let bodyStart = html.range(of: "<body"),
           let bodyTagEnd = html.range(of: ">", range: bodyStart.upperBound..<html.endIndex),
           let bodyEnd = html.range(of: "</body>") {
            return String(html[bodyTagEnd.upperBound..<bodyEnd.lowerBound])
        }

        return html
    }

    nonisolated private func appendMultipartPart(_ body: inout Data,
                                     boundary: String,
                                     contentType: String,
                                     contentDisposition: String,
                                     data: Data) {
        body.append("--\(boundary)\r\n".data(using: .utf8)!)
        body.append("Content-Disposition: \(contentDisposition)\r\n".data(using: .utf8)!)
        body.append("Content-Type: \(contentType)\r\n\r\n".data(using: .utf8)!)
        body.append(data)
        body.append("\r\n".data(using: .utf8)!)
    }

    nonisolated private func escapeHTML(_ value: String) -> String {
        var escaped = value
        escaped = escaped.replacingOccurrences(of: "&", with: "&amp;")
        escaped = escaped.replacingOccurrences(of: "<", with: "&lt;")
        escaped = escaped.replacingOccurrences(of: ">", with: "&gt;")
        escaped = escaped.replacingOccurrences(of: "\"", with: "&quot;")
        escaped = escaped.replacingOccurrences(of: "'", with: "&#39;")
        return escaped
    }
}

