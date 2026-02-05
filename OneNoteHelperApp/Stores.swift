import SwiftUI
import AppKit
import os.log
import Foundation

@MainActor
final class LogStore: ObservableObject {
    static let shared = LogStore()

    @Published var text: String = ""

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
final class SectionStore: ObservableObject {
    struct Section: Identifiable, Hashable {
        let id: String
        let name: String
    }

    static let shared = SectionStore()

    // AppDelegate registers itself here on launch.
    private weak var appDelegate: AppDelegate?

    @Published var sections: [Section] = []
    @Published var isLoading: Bool = false
    @Published var selectedSectionId: String? {
        didSet {
            UserDefaults.standard.set(selectedSectionId, forKey: "TargetSectionId")
            if let id = selectedSectionId, let sec = sections.first(where: { $0.id == id }) {
                UserDefaults.standard.set(sec.name, forKey: "TargetSectionName")
            }
        }
    }

    init() {
        self.selectedSectionId = UserDefaults.standard.string(forKey: "TargetSectionId")
    }

    func register(appDelegate: AppDelegate) {
        self.appDelegate = appDelegate
    }

    func refresh() {
        guard let appDelegate else {
            LogStore.shared.append("ERROR: AppDelegate not registered yet")
            return
        }

        LogStore.shared.append("Refreshing OneNote sections…")
        isLoading = true
        appDelegate.fetchSectionsInDefaultNotebook { result in
            Task { @MainActor in
                self.isLoading = false
                switch result {
                case .success(let sections):
                    LogStore.shared.append("Loaded \(sections.count) section(s)")
                    self.sections = sections
                    // Keep selection if still present, else default to first.
                    if let current = self.selectedSectionId, sections.contains(where: { $0.id == current }) {
                        // ok
                    } else {
                        self.selectedSectionId = sections.first?.id
                    }
                case .failure(let err):
                    LogStore.shared.append("ERROR: failed to load sections: \(err.localizedDescription)")
                }
            }
        }
    }
}
