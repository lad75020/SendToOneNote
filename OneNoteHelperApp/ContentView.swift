import SwiftUI
import AppKit
import Foundation

struct ContentView: View {
    @State private var status: String = "Waiting for print jobs..."
    @State private var showLogs: Bool = false
    @StateObject private var logs = LogStore.shared
    @StateObject private var sections = SectionStore.shared
    @StateObject private var settings = SettingsStore.shared

    private func effectiveClientId() -> String {
        let v = ((Bundle.main.object(forInfoDictionaryKey: "MSALClientId") as? String) ?? UserDefaults.standard.string(forKey: "MSALClientId") ?? "").trimmingCharacters(in: .whitespacesAndNewlines)
        return v
    }

    private func effectiveRedirectUri(clientId: String) -> String {
        let fromPlist = (Bundle.main.object(forInfoDictionaryKey: "MSALRedirectUri") as? String)?.trimmingCharacters(in: .whitespacesAndNewlines) ?? ""
        let fromDefaults = (UserDefaults.standard.string(forKey: "MSALRedirectUri") ?? "").trimmingCharacters(in: .whitespacesAndNewlines)
        if !fromPlist.isEmpty { return fromPlist }
        if !fromDefaults.isEmpty { return fromDefaults }
        let bundleId = Bundle.main.bundleIdentifier ?? ""
        guard !bundleId.isEmpty else { return "" }
        return "msauth.\(bundleId)://auth"
    }
    
    private func effectiveAuthority() -> String {
        let fromPlist = (Bundle.main.object(forInfoDictionaryKey: "MSALAuthority") as? String)?.trimmingCharacters(in: .whitespacesAndNewlines) ?? ""
        let fromDefaults = (UserDefaults.standard.string(forKey: "MSALAuthority") ?? "").trimmingCharacters(in: .whitespacesAndNewlines)
        if !fromPlist.isEmpty { return fromPlist }
        if !fromDefaults.isEmpty { return fromDefaults }
        return "https://login.microsoftonline.com/organizations"
    }

    private func effectiveWatchFolderPath() -> String {
        let fromDefaults = (UserDefaults.standard.string(forKey: "WatchFolderPath") ?? "").trimmingCharacters(in: .whitespacesAndNewlines)
        if !fromDefaults.isEmpty { return fromDefaults }
        let fromSettings = settings.watchFolderPath.trimmingCharacters(in: .whitespacesAndNewlines)
        if !fromSettings.isEmpty { return fromSettings }
        return "/Users/Shared/OneNoteHelper"
    }

    private func isSchemeRegistered(_ scheme: String) -> Bool {
        guard let urlTypes = Bundle.main.object(forInfoDictionaryKey: "CFBundleURLTypes") as? [[String: Any]] else { return false }
        for type in urlTypes {
            if let schemes = type["CFBundleURLSchemes"] as? [String], schemes.contains(scheme) {
                return true
            }
        }
        return false
    }

    var body: some View {
        let cid = effectiveClientId()
        let redirect = effectiveRedirectUri(clientId: cid)
        let scheme = URL(string: redirect)?.scheme ?? ""
        let schemeOK = !scheme.isEmpty && isSchemeRegistered(scheme)
        let authority = effectiveAuthority()
        let watchPath = effectiveWatchFolderPath()
        let watchExists = FileManager.default.fileExists(atPath: watchPath)

        VStack(alignment: .leading, spacing: 12) {
            Text("OneNote Helper")
                .font(.title2)

            Text(status)
                .font(.body)
                .foregroundStyle(.secondary)

            GroupBox("Microsoft Graph / MSAL") {
                VStack(alignment: .leading, spacing: 8) {
                    HStack {
                        Text("MSAL Client ID")
                            .frame(width: 140, alignment: .leading)
                        TextField("xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx", text: $settings.msalClientId)
                    }

                    HStack {
                        Text("MSAL Redirect URI")
                            .frame(width: 140, alignment: .leading)
                        TextField("msauth.<bundleId>://auth", text: $settings.msalRedirectUri)
                    }
                    HStack {
                        Text("MSAL Authority")
                            .frame(width: 140, alignment: .leading)
                        TextField("https://login.microsoftonline.com/organizations", text: $settings.msalAuthority)
                    }

                    HStack {
                        Button("Save") {
                            settings.save()
                            LogStore.shared.append("Saved MSAL settings")
                        }

                        Button("Sign in") {
                            settings.save()
                            AppDelegate.shared?.signIn()
                        }

                        Spacer()
                    }

                    VStack(alignment: .leading, spacing: 4) {
                        HStack(spacing: 8) {
                            Circle().fill(cid.isEmpty ? Color.red : Color.green).frame(width: 8, height: 8)
                            Text("Client ID: \(cid.isEmpty ? "(missing)" : cid)")
                        }
                        .font(.footnote)
                        .foregroundStyle(.secondary)

                        HStack(spacing: 8) {
                            Circle().fill(redirect.isEmpty ? Color.red : Color.green).frame(width: 8, height: 8)
                            Text("Redirect URI: \(redirect.isEmpty ? "(missing)" : redirect)")
                        }
                        .font(.footnote)
                        .foregroundStyle(.secondary)

                        HStack(spacing: 8) {
                            Circle().fill(authority.isEmpty ? Color.red : Color.green).frame(width: 8, height: 8)
                            Text("Authority: \(authority.isEmpty ? "(missing)" : authority)")
                        }
                        .font(.footnote)
                        .foregroundStyle(.secondary)

                        HStack(spacing: 8) {
                            Circle().fill(schemeOK ? Color.green : Color.red).frame(width: 8, height: 8)
                            Text("URL scheme registered: \(scheme.isEmpty ? "(none)" : scheme)\(schemeOK ? "" : " (add to Info.plist)")")
                        }
                        .font(.footnote)
                        .foregroundStyle(.secondary)
                    }

                    Text("Tip: If Sign in shows a web login, complete it once. Afterwards, Refresh can use silent auth. If your app is single-tenant, set MSAL Authority to your tenant endpoint (not /common), e.g. https://login.microsoftonline.com/<tenant-id>.")
                        .font(.footnote)
                        .foregroundStyle(.secondary)
                }
            }

            GroupBox("Watch folder") {
                VStack(alignment: .leading, spacing: 8) {
                    HStack {
                        Text("Root path")
                            .frame(width: 140, alignment: .leading)
                        TextField("/Users/Shared/OneNoteHelper", text: $settings.watchFolderPath)

                        Button("Browse…") {
                            // Ensure we have a visible UI before presenting the open panel.
                            if NSApp.activationPolicy() != .regular {
                                _ = NSApp.setActivationPolicy(.regular)
                            }
                            NSApp.activate(ignoringOtherApps: true)

                            let panel = NSOpenPanel()
                            panel.canChooseFiles = false
                            panel.canChooseDirectories = true
                            panel.allowsMultipleSelection = false
                            panel.prompt = "Choose"

                            let handle: (NSApplication.ModalResponse) -> Void = { resp in
                                if resp == .OK, let url = panel.url {
                                    settings.setWatchFolder(url: url)
                                    LogStore.shared.append("Selected watch folder: \(url.path)")
                                }
                            }

                            if let win = NSApp.keyWindow {
                                panel.beginSheetModal(for: win, completionHandler: handle)
                            } else {
                                handle(panel.runModal())
                            }
                        }
                    }

                    HStack {
                        Button("Save & Restart Watcher") {
                            settings.save()
                            LogStore.shared.append("Saved watch folder: \(settings.watchFolderPath)")
                            AppDelegate.shared?.restartWatchingIncomingFolder()
                        }

                        Spacer()
                    }

                    HStack(spacing: 8) {
                        Circle().fill(watchExists ? Color.green : Color.red).frame(width: 8, height: 8)
                        Text("Effective: \(watchPath)\(watchExists ? "" : " (missing)")")
                    }
                    .font(.footnote)
                    .foregroundStyle(.secondary)

                    Text("This folder must contain Incoming/Processing/Done/Failed. The CUPS backend should drop job-*.pdf + job-*.json into Incoming.")
                        .font(.footnote)
                        .foregroundStyle(.secondary)
                }
            }

            GroupBox("Target section (all notebooks)") {
                HStack {
                    Button(sections.isLoading ? "Loading…" : "Refresh") {
                        sections.refresh()
                    }
                    .disabled(sections.isLoading)

                    Spacer()

                    if let name = UserDefaults.standard.string(forKey: "TargetSectionName") {
                        Text("Selected: \(name)")
                            .foregroundStyle(.secondary)
                    }
                }

                if sections.sections.isEmpty {
                    Text("No sections loaded yet. Click Refresh.")
                        .foregroundStyle(.secondary)
                        .padding(.top, 6)
                } else {
                    Picker("Section", selection: Binding(get: {
                        sections.selectedSectionId ?? ""
                    }, set: { newValue in
                        sections.selectedSectionId = newValue.isEmpty ? nil : newValue
                    })) {
                        ForEach(sections.sections) { s in
                            Text(s.name).tag(s.id)
                        }
                    }
                    .pickerStyle(.menu)
                    .padding(.top, 6)

                    Text("Print jobs will be uploaded to the selected section as a page titled \"Sent To OneNote\".")
                        .font(.footnote)
                        .foregroundStyle(.secondary)
                        .padding(.top, 4)
                }
            }

            HStack {
                Toggle("Show logs", isOn: $showLogs)
                    .toggleStyle(.switch)

                Spacer()

                Button("Clear") {
                    logs.clear()
                }
                .disabled(logs.text.isEmpty)

                Button("Test URL Handler") {
                    if let pdfURL = Bundle.main.url(forResource: "BundledSample", withExtension: "pdf") {
                        var comps = URLComponents()
                        comps.scheme = "onenote-helper"
                        comps.host = "import"
                        comps.queryItems = [
                            URLQueryItem(name: "file", value: pdfURL.path),
                            URLQueryItem(name: "title", value: "Bundled Sample"),
                            URLQueryItem(name: "user", value: "user"),
                            URLQueryItem(name: "job", value: "demo")
                        ]
                        if let url = comps.url {
                            NSApplication.shared.delegate?.application?(NSApplication.shared, open: [url])
                        } else {
                            LogStore.shared.append("ERROR: Failed to build test URL")
                        }
                    } else {
                        LogStore.shared.append("ERROR: BundledSample.pdf not found in bundle resources")
                    }
                }
            }

            if showLogs {
                TextEditor(text: $logs.text)
                    .font(.system(.body, design: .monospaced))
                    .frame(minHeight: 240)
                    .border(Color.secondary.opacity(0.35))
            }
        }
        .padding()
        .frame(minWidth: 620, minHeight: showLogs ? 620 : 340)
    }
}

#Preview {
    ContentView()
}
