# @yome/ppt

PowerPoint editing commands for Yome agents (open / new / save / slides /
slide.add / slide.delete / slide.move / get / read / set / addtext / title /
fmt / export). One of the official skills.

## Layout (spec v0.1, section 4)

```
yome-skill-ppt/
├── yome-skill.json                       manifest (slug / domain / delivery / capabilities)
├── README.md
├── signature/
│   └── ppt.signature.json                LLM-facing command signature (truth)
├── backends/
│   ├── macos/                            Swift module bundled into Yome macOS app
│   ├── ios/                              Swift module bundled into Yome iOS app (read-only subset)
│   ├── node/                             TS backend for Yome CLI / Server
│   └── sandbox/                          TS state machine for hub replays / benchmarks
├── viewer/
│   └── index.html                        single-direction trace renderer
├── cases/                                community-contributed Replays
│   └── example-1/                        new + title + slide.add + save
└── benchmarks/                           officially scored cases
    └── sanity/                           input + fixtures + expected + scorer
```

## Status during v0.1 monorepo phase

The signature in `signature/ppt.signature.json` is byte-identical to
`Server/agent/commands/ppt.signature.json`. A vitest assertion locks the
two files together so they cannot drift; once the CLI installer lands
(Phase 2), `Server/agent/commands/index.ts` reads from this skill-repo
copy directly.

The compress functions live in `backends/sandbox/src/compress.ts` (spec 4.4
location) but are also mirrored at `Server/agent/compress/ppt.ts` because
the Cloudflare Workers runtime can only resolve files under its own
`rootDir`. Same parity test guards them.

The macOS / iOS Swift sources still live in the Yome app target; the
`backends/macos` and `backends/ios` directories here are scaffolds to make
the spec layout faithful and to enable a future `git mv` when the skill is
split into its own repo (spec 8.5).
