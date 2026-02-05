import SwiftUI
import AppKit

@main
struct OneNoteHelperApp: App {
    @NSApplicationDelegateAdaptor(AppDelegate.self) var appDelegate

    var body: some Scene {
        MenuBarExtra("OneNote Helper", systemImage: "note.text") {
            ContentView()
                .tint(.blue)
                .frame(width: 720)

            Divider()

            Button("Quit") {
                NSApp.terminate(nil)
            }
            .keyboardShortcut("q")
        }
        // Window-style popover anchored under the menu bar icon.
        .menuBarExtraStyle(.window)
    }
}
