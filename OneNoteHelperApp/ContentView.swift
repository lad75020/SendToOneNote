import SwiftUI
import AppKit
import Foundation

struct ContentView: View {
    @Environment(\.colorScheme) private var colorScheme

    @State private var showLogs: Bool = false
    @StateObject private var logs = LogStore.shared
    @StateObject private var sections = SectionStore.shared
    @StateObject private var settings = SettingsStore.shared

    /// Controls how the helper builds the OneNote page when importing print jobs.
    /// - image: render pages as PNGs and upload as <img> attachments.
    /// - text: upload extracted text only.
    /// - hybrid: current default (text + embedded images when found; falls back to rendered pages if needed).
    @AppStorage("ImportMode") private var importMode: String = "hybrid"

    // MARK: - Theme

    private var pageBackground: Color {
        // Slightly tinted blue background, adaptive.
        colorScheme == .dark
            ? Color(red: 0.06, green: 0.10, blue: 0.18)
            : Color(red: 0.92, green: 0.96, blue: 1.0)
    }

    private var panelBackground: Color {
        // Fallback solid color if material isn't desired.
        colorScheme == .dark
            ? Color(red: 0.09, green: 0.14, blue: 0.24)
            : Color.white.opacity(0.90)
    }

    private var panelShadow: Color {
        colorScheme == .dark ? Color.black.opacity(0.35) : Color.black.opacity(0.08)
    }

    private var primaryText: Color {
        colorScheme == .dark ? Color(red: 0.92, green: 0.96, blue: 1.0) : .black
    }

    private var secondaryText: Color {
        colorScheme == .dark ? Color.white.opacity(0.75) : .black.opacity(0.6)
    }

    private var accentGray: Color {
        colorScheme == .dark ? Color.white.opacity(0.22) : Color.black.opacity(0.18)
    }

    @ViewBuilder
    private func panel<Content: View>(title: String, @ViewBuilder content: () -> Content) -> some View {
        VStack(alignment: .leading, spacing: 12) {
            Text(title)
                .font(.headline)
                .foregroundStyle(primaryText)

            content()
        }
        .padding(14)
        .background(
            RoundedRectangle(cornerRadius: 14, style: .continuous)
                .fill(.regularMaterial)
        )
        .overlay(
            RoundedRectangle(cornerRadius: 14, style: .continuous)
                .stroke(accentGray, lineWidth: 1)
        )
        .shadow(color: panelShadow, radius: 10, x: 0, y: 4)
    }

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

        VStack(alignment: .leading, spacing: 14) {
            HStack(alignment: .firstTextBaseline, spacing: 10) {
                Image(systemName: "note.text")
                    .font(.title2)
                    .foregroundStyle(Color.blue)
                    .accessibilityHidden(true)

                Text("OneNote Helper")
                    .font(.title2)
                    .foregroundStyle(primaryText)

                Spacer()

                HStack(spacing: 8) {
                    ProgressView()
                        .controlSize(.small)
                        .opacity(logs.activityState == .waiting ? 0 : 1)

                    Text(logs.activityState == .waiting ? "Waiting for print jobs…" : "\(logs.activityState.label)…")
                        .font(.callout)
                        .foregroundStyle(secondaryText)
                        .help("Current activity: \(logs.activityState.label)")
                }
            }

            panel(title: "Microsoft Graph / MSAL") {
                VStack(alignment: .leading, spacing: 10) {
                    HStack {
                        Text("MSAL Client ID")
                            .frame(width: 140, alignment: .leading)
                            .foregroundStyle(primaryText)
                        TextField("xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx", text: $settings.msalClientId)
                            .textFieldStyle(.roundedBorder)
                            .help("Azure AD application (client) ID used for MSAL sign-in.")
                    }

                    HStack {
                        Text("MSAL Redirect URI")
                            .frame(width: 140, alignment: .leading)
                            .foregroundStyle(primaryText)
                        TextField("msauth.<bundleId>://auth", text: $settings.msalRedirectUri)
                            .textFieldStyle(.roundedBorder)
                            .help("Redirect URI configured in Azure for this app (msauth.<bundle-id>://auth).")
                    }
                    HStack {
                        Text("MSAL Authority")
                            .frame(width: 140, alignment: .leading)
                            .foregroundStyle(primaryText)
                        TextField("https://login.microsoftonline.com/organizations", text: $settings.msalAuthority)
                            .textFieldStyle(.roundedBorder)
                            .help("Authority used for sign-in. Use /organizations or your tenant URL.")
                    }

                    HStack(spacing: 10) {
                        Button("Save") {
                            settings.save()
                            LogStore.shared.append("Saved MSAL settings")
                        }
                        .buttonStyle(.bordered)
                        .controlSize(.regular)
                        .help("Save MSAL settings (Client ID / Redirect URI / Authority).")

                        Button("Sign in") {
                            settings.save()
                            AppDelegate.shared?.signIn()
                        }
                        .buttonStyle(.borderedProminent)
                        .controlSize(.regular)
                        .help("Authenticate with Microsoft so the app can call the Graph API.")

                        Spacer()
                    }

                    HStack(spacing: 10) {
                        Circle().fill(cid.isEmpty ? Color.red : Color.blue).frame(width: 8, height: 8)
                            .help(cid.isEmpty ? "Missing MSAL Client ID" : "MSAL Client ID configured")
                        Circle().fill(redirect.isEmpty ? Color.red : Color.blue).frame(width: 8, height: 8)
                            .help(redirect.isEmpty ? "Missing Redirect URI" : "Redirect URI configured")
                        Circle().fill(authority.isEmpty ? Color.red : Color.blue).frame(width: 8, height: 8)
                            .help(authority.isEmpty ? "Missing Authority" : "Authority configured")
                        Circle().fill(schemeOK ? Color.blue : Color.red).frame(width: 8, height: 8)
                            .help(schemeOK ? "URL scheme registered" : "URL scheme not registered (Info.plist)")

                        Spacer()
                    }
                    .padding(.top, 2)

                }
            }

            panel(title: "Watch folder") {
                VStack(alignment: .leading, spacing: 10) {
                    HStack {
                        Text("Root path")
                            .frame(width: 140, alignment: .leading)
                            .foregroundStyle(primaryText)
                        TextField("/Users/Shared/OneNoteHelper", text: $settings.watchFolderPath)
                            .textFieldStyle(.roundedBorder)
                            .help("Root folder that contains Incoming/Processing/Done/Failed.")

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
                        .buttonStyle(.bordered)
                        .help("Pick the watch folder on disk.")
                    }

                    HStack {
                        Button("Save & Restart Watcher") {
                            settings.save()
                            LogStore.shared.append("Saved watch folder: \(settings.watchFolderPath)")
                            AppDelegate.shared?.restartWatchingIncomingFolder()
                        }
                        .buttonStyle(.borderedProminent)
                        .help("Save the folder path and restart the incoming print-job watcher.")

                        Spacer()

                        // Compact status indicator
                        Label(watchExists ? "Ready" : "Missing", systemImage: watchExists ? "checkmark.circle.fill" : "exclamationmark.triangle.fill")
                            .font(.footnote)
                            .foregroundStyle(watchExists ? Color.blue : Color.red)
                            .help("Effective watch folder: \(watchPath)")
                    }
                }
            }

            panel(title: "Import mode") {
                VStack(alignment: .leading, spacing: 10) {
                    Picker("Mode", selection: $importMode) {
                        Text("Image").tag("image")
                        Text("Text").tag("text")
                        Text("Hybrid").tag("hybrid")
                    }
                    .pickerStyle(.segmented)
                    .controlSize(.regular)
                    .help("Choose how pages are generated in OneNote: images only, text only, or hybrid.")
                }
            }

            panel(title: "Target section") {
                VStack(alignment: .leading, spacing: 10) {
                    HStack {
                        Button(sections.isLoading ? "Loading…" : "Refresh") {
                            sections.refresh()
                        }
                        .disabled(sections.isLoading)
                        .buttonStyle(.bordered)
                        .help("Load all OneNote sections and refresh the list.")

                        Spacer()

                        if let name = UserDefaults.standard.string(forKey: "TargetSectionName") {
                            Text(name)
                                .foregroundStyle(secondaryText)
                                .help("Currently selected section.")
                        }
                    }

                    Picker("Section", selection: Binding(get: {
                        sections.selectedSectionId ?? ""
                    }, set: { newValue in
                        sections.selectedSectionId = newValue.isEmpty ? nil : newValue
                    })) {
                        Text("(None)").tag("")
                        ForEach(sections.sections) { s in
                            Text(s.name).tag(s.id)
                        }
                    }
                    .pickerStyle(.menu)
                    .help("Where new OneNote pages will be created.")
                }
            }

            HStack {
                Toggle("Show logs", isOn: $showLogs)
                    .toggleStyle(.switch)
                    .help("Show or hide detailed logs.")

                Spacer()

                Button("Clear") {
                    logs.clear()
                }
                .disabled(logs.text.isEmpty)
                .help("Clear the log view.")

            }

            if showLogs {
                TextEditor(text: $logs.text)
                    .font(.system(.body, design: .monospaced))
                    .frame(minHeight: 240)
                    .padding(8)
                    .background(
                        RoundedRectangle(cornerRadius: 10, style: .continuous)
                            .fill(colorScheme == .dark ? Color.black.opacity(0.18) : Color.white.opacity(0.75))
                    )
                    .overlay(
                        RoundedRectangle(cornerRadius: 10, style: .continuous)
                            .stroke(accentGray, lineWidth: 1)
                    )
            }
        }
        .padding(16)
        .background(pageBackground)
        .frame(minWidth: 640, minHeight: showLogs ? 640 : 360)
    }
}

#Preview {
    ContentView()
}
