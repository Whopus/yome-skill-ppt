// swift-tools-version:5.9
//
// @yome/ppt — macOS backend (Swift module compiled into the Yome macOS app).
// During the v0.1 monorepo phase, the actual implementation lives in the
// main app at Yome/Shared/* and Yome/macOS/*. This Package.swift exists so
// the skill repo layout matches the spec 4 standard structure; the bundled
// integration is via the app's own Xcode project, not SwiftPM.
//
// TODO(spec-v0.1): once the official ppt skill is split out of the monorepo
// (spec 8.5), move the corresponding Swift sources into Sources/PptBackend/.

import PackageDescription

let package = Package(
    name: "PptBackend",
    platforms: [.macOS(.v13)],
    products: [
        .library(name: "PptBackend", targets: ["PptBackend"])
    ],
    targets: [
        .target(name: "PptBackend", path: "Sources/PptBackend")
    ]
)
