[app]
title = 作业票查询
package.name = zhuyepiao
package.domain = org.hse
source.dir = .
source.include_exts = py,png,jpg,kv,atlas,ttf,json
source.exclude_dirs = .buildozer, bin, __pycache__
version = 3.0

# 依赖包：requests 和 openpyxl 需要的依赖都已经包含
requirements = python3,kivy==2.3.0,requests,urllib3,charset-normalizer,idna,certifi,openpyxl,et-xmlfile,android

orientation = portrait
fullscreen = 0
android.permissions = INTERNET,READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE
android.api = 33
android.minapi = 21
android.ndk = 25b
android.ndk_api = 21
android.archs = arm64-v8a, armeabi-v7a
android.accept_sdk_license = True
android.wakelock = 1

[buildozer]
log_level = 2
warn_on_root = 1