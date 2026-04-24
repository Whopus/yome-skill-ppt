// swift-tools-version:5.9
//
// @yome/ppt — iOS backend (Swift module compiled into the Yome iOS app).
// Most ppt operations are not available on iOS (no Keynote AppleScript
// surface), so this backend exposes only the read-only commands. The
// implementation lives in the Yome iOS app target during the v0.1
// monorepo phase.
//
// TODO(spec-v0.1): once the official ppt skill is split out (spec 8.5),
// move iOS-specific Swift sources into Sources/PptBackend/.

import PackageDescription

let package = Package(
    name: "PptBackend",
    platforms: [.iOS(.v17)],
    products: [
        .library(name: "PptBackend", targets: ["PptBackend"])
    ],
    targets: [
        .target(name: "PptBackend", path: "Sources/PptBackend")
    ]
)
