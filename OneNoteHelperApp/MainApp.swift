import SwiftUI
import AppKit

@main
struct OneNoteHelperApp: App {
    @NSApplicationDelegateAdaptor(AppDelegate.self) var appDelegate

    var body: some Scene {
        WindowGroup {
            ContentView()
                .tint(.blue)
        }
    }
}
