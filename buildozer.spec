[app]
title = 対話分析
package.name = analysedialogue
package.domain = org.kivy

source.dir = .
source.include_exts = py
source.include_patterns = dialogue_core.py,analyze_dialogue_android.py

version = 1.0

# Pure-Python packages only — pydantic-core (Rust) is intentionally excluded.
# anthropic SDK is replaced by raw requests calls in analyze_dialogue_android.py.
requirements = python3,kivy==2.3.0,requests,openpyxl,et_xmlfile,plyer,certifi,charset-normalizer,idna,urllib3

orientation = portrait
fullscreen = 0

android.permissions = android.permission.INTERNET,android.permission.READ_EXTERNAL_STORAGE,android.permission.WRITE_EXTERNAL_STORAGE
android.api = 34
android.minapi = 26
android.ndk = 25b
android.accept_sdk_license = True

# Entry point
entrypoint = analyze_dialogue_android
entrypoint_class = DialogueAnalysisApp

[buildozer]
log_level = 2
warn_on_root = 1
