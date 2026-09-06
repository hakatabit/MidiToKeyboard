# MidiToKeyboard

## 概要

MidiToKeyboardは、MIDIキーボードやパッドからの入力をPCのキーボード入力へ変換するWindows向けアプリケーションです。MIDIのNoteOn / NoteOffを、対象アプリケーションへKeyDown / KeyUpとして送出します。

現在はコンソールUIで動作します。内部はUi / Application / Domain / Infrastructureに分割されており、将来タスクトレイUIへ移行するときに、MIDI変換や外部依存処理への変更を小さくできる構成を目指しています。

## 目的

MIDIキーボードやパッドを使って、キーボード入力に対応した音楽ゲームをプレイできるようにすることを目的としています。

## 主な機能

- MIDI入力デバイスの一覧表示と選択
- 入力送信モードの選択
  - VirtualKey
  - Scancode
- `mappings.json`に定義されたプロファイルの一覧表示と選択
- MIDIノート番号とキー文字のマッピング
- NoteOnに応じたKeyDown、NoteOffに応じたKeyUpの送出
- 同じキーに複数ノートが割り当てられた場合の参照カウント管理
- MIDI ON / OFFのコンソール表示
- Windows `SendInput`失敗時の警告表示

## 実行時の基本フロー

1. MIDI入力デバイスを選択します。
2. 入力送信モードを選択します。
   - `[1]` VirtualKey
   - `[2]` Scancode
   - `2`以外の入力はVirtualKeyとして扱う
3. `mappings.json`から読み込まれたプロファイルを選択します。
4. MIDI入力の待ち受けを開始します。
5. MIDI入力に応じてキー入力を送出します。
6. コンソールで任意のキーを押すと停止・終了します。

## `mappings.json`の設定

`mappings.json`は実行ファイルと同じディレクトリから読み込まれます。Visual Studioでビルドすると出力ディレクトリへコピーされます。

`sets`には複数のプロファイルを定義できます。`name`が選択画面に表示され、`mappings`にMIDIノート番号とキー文字の対応を記述します。

```json
{
  "sets": [
    {
      "name": "Default",
      "mappings": {
        "52": "v",
        "53": "s",
        "55": "d",
        "57": "f",
        "59": "g",
        "65": "h",
        "67": "j",
        "69": "k",
        "71": "l",
        "72": "n"
      }
    }
  ]
}
```

- 1つのノート番号には1つのキー文字を割り当てます。
- キー文字列の先頭文字がマッピングに使用されます。
- 選択可能なプロファイルがない場合は、メッセージを表示して終了します。
- ファイルが存在しない、またはJSONを読み込めない場合は、エラーメッセージを表示して終了します。
- 代替マッピングは自動生成しません。

ノート名とノート番号の対応は`NoteNumberList.md`を参照してください。

## 開発・動作環境

- OS: Windows
- IDE: Visual Studio 2022
- 対象フレームワーク: .NET Framework 4.7.2
- プロジェクト形式: 旧形式csproj
- JSON読み込み: System.Text.Json (NuGet)
- MIDI入力: [Melanchall.DryWetMidi](https://github.com/melanchall/drywetmidi)
- キー送出: Windows API `SendInput`

既存の確認実績はWindows 11上のものです。Windows 10など、その他のWindowsバージョンで使用する場合は実機確認してください。

## ビルド方法

1. Visual Studioで`MidiToKeyboard.sln`を開きます。
2. NuGetパッケージを復元します。
3. DebugまたはRelease構成でビルドします。
4. 出力ディレクトリに`mappings.json`がコピーされていることを確認します。

詳細な確認項目は`TESTING.md`を参照してください。

## 注意事項

- ビルドと実行はWindows上のVisual Studioを前提としています。WSL上でのビルドは前提にしていません。
- .NET Framework 4.7.2と旧形式csprojの互換性を維持します。
- C# 10以上の構文は使用しない方針です。
- namespaceには従来のブロック形式を使用し、ファイルスコープnamespaceは使用しません。
- キー入力先のウィンドウにフォーカスが必要です。
- 対象アプリケーションとの権限差やOSの制限により、`SendInput`が届かない場合があります。
- 現在はタスクトレイUIに対応していません。

## 動作確認タイトル

既存のREADMEには、以下のPC版タイトルでの確認実績が記録されています。

- DJMAX RESPECT V（Xbox / Steam）
- MUSYNX（Xbox / Steam）
- EZ2ON（Steam）

XboxはXbox Play AnywhereによるPC版です。DJMAX RESPECT VではScancodeモードを使用します。他のアプリケーションでの動作は個別に確認してください。

## 今後の予定

- コンソールUIからタスクトレイUIへの移行
- UI差し替え時のDomain / Infrastructureへの変更を最小限に保つ
- 必要に応じた自動テストの追加

## ライセンス

このプロジェクトはMITライセンスの下で公開されています。詳細は`LICENSE`を参照してください。

依存ライブラリのライセンスは、それぞれの作者に帰属します。
