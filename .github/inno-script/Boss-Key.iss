name: Build and Release Installer

on:
  push:
    tags:
      - "v*"

permissions:
  contents: write

jobs:
  build-installer:
    runs-on: windows-latest

    steps:
      - name: Checkout
        uses: actions/checkout@v4

      # 如果你还有前置构建（例如生成 software/Boss-Key.exe），在这里加
      # - name: Build app
      #   run: ...

      - name: Install Inno Setup
        shell: pwsh
        run: |
          choco install innosetup --no-progress -y

      - name: Locate ISCC
        id: iscc
        shell: pwsh
        run: |
          $iscc = (Get-Command ISCC.exe -ErrorAction Stop).Source
          "ISCC=$iscc" | Out-File -FilePath $env:GITHUB_ENV -Encoding utf8 -Append
          Write-Host "Found ISCC at: $iscc"

      - name: Verify required files
        shell: pwsh
        run: |
          if (!(Test-Path ".\.github\inno-script\Boss-Key.iss")) { throw "Missing .iss file" }
          if (!(Test-Path ".\.github\inno-script\Languages\ChineseSimplified.isl")) { throw "Missing ChineseSimplified.isl" }
          if (!(Test-Path ".\.github\inno-script\software\Boss-Key.exe")) { throw "Missing software\Boss-Key.exe" }

      - name: Compile Installer
        shell: pwsh
        run: |
          $version = "${{ github.ref_name }}"   # 例如 v2.1.1
          & "$env:ISCC" "/DMyAppVersion=$version" ".\.github\inno-script\Boss-Key.iss"

      - name: Upload installer artifact
        uses: actions/upload-artifact@v4
        with:
          name: Boss-Key-Installer-${{ github.ref_name }}
          path: ./.github/inno-script/output/*.exe

      - name: Create GitHub Release
        uses: softprops/action-gh-release@v1
        with:
          files: ./.github/inno-script/output/*.exe
          tag_name: ${{ github.ref_name }}
          name: ${{ github.ref_name }}
          generate_release_notes: true
