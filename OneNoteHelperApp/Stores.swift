import SwiftUI
import AppKit
import os.log
import Foundation

@MainActor
final class LogStore: ObservableObject {
    static let shared = LogStore()

    enum ActivityState: String {
        case waiting
        case processing
        case uploading

        var label: String {
            switch self {
            case .waiting: return "Waiting"
            case .processing: return "Processing"
            case .uploading: return "Uploading"
            }
        }
    }

    @Published var text: String = ""
    @Published var activityState: ActivityState = .waiting

    private let df: DateFormatter = {
        let d = DateFormatter()
        d.dateFormat = "HH:mm:ss"
        return d
    }()

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

    func append(_ line: String) {
        let stamped = "[\(df.string(from: Date()))] \(line)"
        if text.isEmpty {
            text = stamped
        } else {
            text += "\n" + stamped
        }

        // Prevent unbounded memory growth when the watcher is chatty.
        // Keep roughly the last ~4000 lines.
        let maxLines = 4000
        let lines = text.split(separator: "\n", omittingEmptySubsequences: false)
        if lines.count > maxLines {
            text = lines.suffix(maxLines).joined(separator: "\n")
        }
    }

    func clear() {
        text = ""
    }
}

@MainActor
final class SettingsStore: ObservableObject {
    static let shared = SettingsStore()

    @Published var msalClientId: String
    @Published var msalRedirectUri: String
    @Published var msalAuthority: String

    /// Root folder that contains Incoming/Processing/Done/Failed.
    /// Default: /Users/Shared/OneNoteHelper
    @Published var watchFolderPath: String

    /// Security-scoped bookmark for sandboxed builds.
    /// If present, it is used to access the watch folder.
    @Published var watchFolderBookmarkData: Data?

    // Persist for future runs
    private let clientIdKey = "MSALClientId"
    private let redirectKey = "MSALRedirectUri"
    private let authorityKey = "MSALAuthority"
    private let watchFolderKey = "WatchFolderPath"
    private let watchFolderBookmarkKey = "WatchFolderBookmark"

    init() {
        self.msalClientId = UserDefaults.standard.string(forKey: clientIdKey) ?? ""
        self.msalRedirectUri = UserDefaults.standard.string(forKey: redirectKey) ?? ""
        self.msalAuthority = UserDefaults.standard.string(forKey: authorityKey) ?? ""
        self.watchFolderPath = UserDefaults.standard.string(forKey: watchFolderKey) ?? "/Users/Shared/OneNoteHelper"
        self.watchFolderBookmarkData = UserDefaults.standard.data(forKey: watchFolderBookmarkKey)
    }

    func save() {
        UserDefaults.standard.set(msalClientId.trimmingCharacters(in: .whitespacesAndNewlines), forKey: clientIdKey)
        UserDefaults.standard.set(msalRedirectUri.trimmingCharacters(in: .whitespacesAndNewlines), forKey: redirectKey)
        UserDefaults.standard.set(msalAuthority.trimmingCharacters(in: .whitespacesAndNewlines), forKey: authorityKey)
        UserDefaults.standard.set(watchFolderPath.trimmingCharacters(in: .whitespacesAndNewlines), forKey: watchFolderKey)
        if let watchFolderBookmarkData {
            UserDefaults.standard.set(watchFolderBookmarkData, forKey: watchFolderBookmarkKey)
        }
    }

    func setWatchFolder(url: URL) {
        watchFolderPath = url.path
        // Create security-scoped bookmark when sandboxed; harmless otherwise.
        if let data = try? url.bookmarkData(options: [.withSecurityScope], includingResourceValuesForKeys: nil, relativeTo: nil) {
            watchFolderBookmarkData = data
            UserDefaults.standard.set(data, forKey: watchFolderBookmarkKey)
        }
    }
}

@MainActor
final class OneNoteTargetStore: ObservableObject {
    struct Notebook: Identifiable, Hashable { let id: String; let name: String }
    struct Section: Identifiable, Hashable { let id: String; let name: String }
    struct Page: Identifiable, Hashable { let id: String; let title: String }

    static let shared = OneNoteTargetStore()

    private weak var appDelegate: AppDelegate?

    @Published var notebooks: [Notebook] = []
    @Published var sections: [Section] = []
    @Published var pages: [Page] = []

    @Published var isLoading: Bool = false

    @Published var selectedNotebookId: String? {
        didSet {
            UserDefaults.standard.set(selectedNotebookId, forKey: "TargetNotebookId")
            if let id = selectedNotebookId, let nb = notebooks.first(where: { $0.id == id }) {
                UserDefaults.standard.set(nb.name, forKey: "TargetNotebookName")
            }
            // Changing notebook invalidates section/page selections.
            if oldValue != selectedNotebookId {
                selectedSectionId = nil
                selectedPageId = nil
                sections = []
                pages = []
                refreshSectionsForSelectedNotebook()
            }
        }
    }

    @Published var selectedSectionId: String? {
        didSet {
            UserDefaults.standard.set(selectedSectionId, forKey: "TargetSectionId")
            if let id = selectedSectionId, let sec = sections.first(where: { $0.id == id }) {
                UserDefaults.standard.set(sec.name, forKey: "TargetSectionName")
            }
            // Changing section invalidates page selection.
            if oldValue != selectedSectionId {
                selectedPageId = nil
                pages = []
                refreshPagesForSelectedSection()
            }
        }
    }

    /// nil means "None" (create new pages). If set, append to that page.
    @Published var selectedPageId: String? {
        didSet {
            UserDefaults.standard.set(selectedPageId, forKey: "TargetPageId")
            if let id = selectedPageId, let page = pages.first(where: { $0.id == id }) {
                UserDefaults.standard.set(page.title, forKey: "TargetPageTitle")
            }
        }
    }

    var selectionSummary: String? {
        let nb = UserDefaults.standard.string(forKey: "TargetNotebookName")
        let sec = UserDefaults.standard.string(forKey: "TargetSectionName")
        let page = UserDefaults.standard.string(forKey: "TargetPageTitle")
        if let nb, !nb.isEmpty, let sec, !sec.isEmpty {
            if let pid = UserDefaults.standard.string(forKey: "TargetPageId"), !pid.isEmpty {
                return "\(nb) / \(sec) / \(page ?? "(page)")"
            }
            return "\(nb) / \(sec) / (new pages)"
        }
        return nil
    }

    init() {
        self.selectedNotebookId = UserDefaults.standard.string(forKey: "TargetNotebookId")
        self.selectedSectionId = UserDefaults.standard.string(forKey: "TargetSectionId")
        self.selectedPageId = UserDefaults.standard.string(forKey: "TargetPageId")
    }

    func register(appDelegate: AppDelegate) {
        self.appDelegate = appDelegate
    }

    func refreshAll() {
        refreshNotebooks()
    }

    func refreshNotebooks() {
        guard let appDelegate else {
            LogStore.shared.append("ERROR: AppDelegate not registered yet")
            return
        }
        LogStore.shared.append("Refreshing OneNote notebooks…")
        isLoading = true
        appDelegate.fetchNotebooks { result in
            Task { @MainActor in
                switch result {
                case .success(let items):
                    self.notebooks = items
                    // Keep existing selection if valid, else default to first.
                    if let current = self.selectedNotebookId, items.contains(where: { $0.id == current }) {
                        // ok
                    } else {
                        self.selectedNotebookId = items.first?.id
                    }
                case .failure(let err):
                    LogStore.shared.append("ERROR: failed to load notebooks: \(err.localizedDescription)")
                }
                self.isLoading = false
            }
        }
    }

    private func refreshSectionsForSelectedNotebook() {
        guard let appDelegate else { return }
        guard let nbId = selectedNotebookId, !nbId.isEmpty else { return }

        LogStore.shared.append("Refreshing sections for notebook…")
        isLoading = true
        appDelegate.fetchSections(notebookId: nbId) { result in
            Task { @MainActor in
                switch result {
                case .success(let items):
                    self.sections = items
                    // Default to first section if none.
                    if let current = self.selectedSectionId, items.contains(where: { $0.id == current }) {
                        // ok
                    } else {
                        self.selectedSectionId = items.first?.id
                    }
                case .failure(let err):
                    LogStore.shared.append("ERROR: failed to load sections: \(err.localizedDescription)")
                }
                self.isLoading = false
            }
        }
    }

    private func refreshPagesForSelectedSection() {
        guard let appDelegate else { return }
        guard let secId = selectedSectionId, !secId.isEmpty else { return }

        LogStore.shared.append("Refreshing pages for section…")
        isLoading = true
        appDelegate.fetchPages(sectionId: secId) { result in
            Task { @MainActor in
                switch result {
                case .success(let items):
                    self.pages = items
                    // Always default page selection to "None".
                    self.selectedPageId = nil
                case .failure(let err):
                    LogStore.shared.append("ERROR: failed to load pages: \(err.localizedDescription)")
                }
                self.isLoading = false
            }
        }
    }
}
