// PptBackend.swift — iOS bundled backend for @yome/ppt.
// iOS exposes only a read-only subset; mutation is delegated to a peer
// macOS device via @-mention routing.

import Foundation

public enum PptBackend {
    public static let domain = "ppt"
    public static let signatureRange = ">=1.0.0 <2.0.0"
    /// Commands this iOS backend implements. The runtime filters the prompt
    /// to only show these so the LLM does not call mutations on iOS.
    public static let supportedActions: Set<String> = ["files", "slides", "get", "read"]
}
