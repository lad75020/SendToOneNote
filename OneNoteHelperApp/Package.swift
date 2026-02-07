// swift-tools-version: 6.2
import PackageDescription

let package = Package(
    name: "OneNoteHelper",
    platforms: [
        .macOS(.v13)
    ],
    products: [
        .executable(name: "OneNoteHelperApp", targets: ["OneNoteHelperApp"]),
        .executable(name: "onenote_backend", targets: ["OneNoteCUPSBackend"])
    ],
    dependencies: [
        // Microsoft Authentication Library for Apple platforms (MSAL)
        .package(url: "https://github.com/AzureAD/microsoft-authentication-library-for-objc.git", from: "1.4.0")
    ],
    targets: [
        .executableTarget(
            name: "OneNoteHelperApp",
            dependencies: [
                .product(name: "MSAL", package: "microsoft-authentication-library-for-objc")
            ],
            path: ".",
            sources: [
                "OneNoteHelperApp.swift"
            ],
            resources: [
                // Optional: bundle Ghostscript `gs` binary to enable PS->PDF conversion under App Sandbox.
                // See Resources/ghostscript/README.md
                .copy("Resources/ghostscript")
            ]
        ),
        .executableTarget(
            name: "OneNoteCUPSBackend",
            path: "Sources/OneNoteCUPSBackend",
            sources: ["onenote_backend.c"],
            linkerSettings: [
                // Link against libcups which provides CUPS backend symbols.
                .linkedLibrary("cups")
            ]
        )
    ]
)

