# TESTING.md

## 目的

この文書は、UI / Application / Domain / Infrastructure への分割後に、Windows の Visual Studio 上で確認すべきビルド観点と手動動作確認観点を整理するものです。

現時点ではテストコードの追加ではなく、既存挙動を変えずにリファクタ後の状態を確認するための手順と観点を記録します。

## 前提環境

- Windows 上の Visual Studio で確認する。
- 対象フレームワークは .NET Framework 4.7.2。
- プロジェクト形式は旧形式 csproj。
- WSL 上でのビルドや実行確認は前提にしない。
- MIDI 入力確認には、Windows から認識できる MIDI 入力デバイスを用意する。
- キー送出確認には、入力を受け取れるテキストエディタや対象アプリを用意する。
- 旧ルートの `MidiToKeyboard.cs` には既存の console 実行経路が残っている。
- 新しい `Application.MidiToKeyboardApplication` と `Ui.ConsoleUi` は追加済みだが、既存実行経路へ全面移行済みとは限らない。

## ビルド確認

1. Visual Studio で `MidiToKeyboard.sln` を開く。
2. NuGet パッケージが復元されていることを確認する。
3. Debug 構成でビルドする。
4. Release 構成でビルドする。
5. .NET Framework 4.7.2 を対象としてビルドできることを確認する。
6. 旧形式 csproj の `Compile Include` に `src` 配下の新規ファイルが含まれていることを確認する。
7. ファイル名変更後の古い `Compile Include` や、存在しないファイルへの参照が残っていないことを確認する。
8. `mappings.json` が出力先へコピーされる設定になっていることを確認する。
9. ビルド警告や名前解決エラーが増えていないことを確認する。

## 起動確認

1. Visual Studio から Debug 実行する。
2. 既存の console 実行経路が起動することを確認する。
3. MIDI 入力デバイス一覧が表示されることを確認する。
4. 既存の表示内容が大きく変わっていないことを確認する。
5. 起動直後に例外で即終了しないことを確認する。
6. MIDI デバイスが存在しない環境では、「デバイスが見つからない」旨の表示で終了することを確認する。

## MIDI入力確認

1. Windows で対象 MIDI デバイスが認識されていることを確認する。
2. アプリ起動後、MIDI デバイス一覧が取得できることを確認する。
3. 対象 MIDI デバイスを番号で選択できることを確認する。
4. デバイス選択後、イベント受信待ち状態になることを確認する。
5. MIDI デバイスから NoteOn を送信し、console に NoteOn 相当のログが表示されることを確認する。
6. MIDI デバイスから NoteOff を送信し、console に NoteOff 相当のログが表示されることを確認する。
7. Velocity 0 の NoteOn が NoteOff 相当として扱われることを確認する。
8. ControlChange など、既存経路で処理対象外のイベントを送っても、想定外の例外で終了しないことを確認する。

## キー送出確認

1. 既存経路でキー入力が送出されることを確認する。
2. 入力送信モードで `VirtualKey` を選択した場合、対象アプリにキー入力が届くことを確認する。
3. 入力送信モードで `Scancode` を選択した場合、対象アプリにキー入力が届くことを確認する。
4. NoteOn で KeyDown 相当が発生することを確認する。
5. NoteOff で KeyUp 相当が発生することを確認する。
6. Velocity 0 の NoteOn でも KeyUp 相当が発生することを確認する。
7. 同じキーに複数ノートが割り当てられている場合、参照カウント挙動が壊れていないことを確認する。
   - 片方のノートを離しても、もう片方のノートが押下中なら KeyUp 相当が早く送られないこと。
   - 両方のノートを離した時点で KeyUp 相当が送られること。
8. 同じノートの重複 NoteOn / 重複 NoteOff で、不要な KeyDown / KeyUp が連続送出されないことを確認する。
9. SendInput 失敗時の console 表示挙動が維持されていることを確認する。
10. `WindowsKeyOutput.WarningOccurred` が購読されている経路では、SendInput 失敗通知が console に表示されることを確認する。
11. `WindowsKeyOutput.WarningOccurred` が未購読の場合、失敗通知は表示されないが例外にはならないことを確認する。

## mappings.json / Profile 読み込み確認

1. `mappings.json` が存在する場合に読み込めることを確認する。
2. `mappings.json` が出力先ディレクトリにコピーされていることを確認する。
3. 複数のマッピングセットがある場合、一覧表示または一覧取得できることを確認する。
4. 選択したマッピングセットが適用されることを確認する。
5. 選択したセットのノート番号とキー文字の対応でキー送出されることを確認する。
6. 存在しないマッピング名を指定した場合の挙動を確認する。
   - 既存 console 経路では、番号選択による挙動を確認する。
   - `JsonProfileRepository.Load` を使う経路では、存在しない名前で例外になることを確認する。
7. `mappings.json` の JSON 構造を変更していないことを確認する。
8. 空の mappings、空のセット、解釈できないノート番号が含まれる場合に、想定外の例外で終了しないか確認する。

## レイヤ分割後の確認観点

1. Domain が Console / DryWetMIDI / Windows API / JSON に依存していないことを確認する。
2. Domain が Application / Infrastructure / Ui の namespace に依存していないことを確認する。
3. Application が Infrastructure の具象クラスに直接依存していないことを確認する。
4. Application が外部機能を `IMidiInput` / `IKeyOutput` / `IProfileRepository` 経由で扱っていることを確認する。
5. Infrastructure が UI に依存していないことを確認する。
6. Infrastructure が DryWetMIDI、Win32 SendInput、JSON 読み込みなどの外部依存を閉じ込めていることを確認する。
7. Ui が Console 入出力を担当していることを確認する。
8. Ui が `KeyPressState` や `MidiTranslator` の内部状態管理へ過剰に踏み込んでいないことを確認する。
9. `global::` が名前衝突回避に必要な箇所だけで使われていることを確認する。
10. `namespace MidiToKeyboard` と `class MidiToKeyboard` の衝突が悪化していないことを確認する。

## 既知の未対応事項

- 旧ルート `MidiToKeyboard.cs` には既存処理が残っている。後続タスクで `Program` / `ConsoleUi` / Application 経由の構成へ整理予定。
- 新しい `Application.MidiToKeyboardApplication` / `Ui.ConsoleUi` は追加済みだが、既存実行経路へ全面移行済みではない可能性がある。Visual Studio 実行時にどの経路が起動しているか確認する。

## トラブルシュート

### 新規ファイルがビルド対象に含まれていない場合

- `MidiToKeyboard.csproj` の `Compile Include` に対象ファイルが含まれているか確認する。
- Visual Studio のソリューション エクスプローラーで対象ファイルがプロジェクト配下に表示されているか確認する。
- ファイルを追加しただけで csproj に反映されていない場合は、プロジェクトに既存項目として追加する。

### namespace と class の名前衝突が疑われる場合

- `namespace MidiToKeyboard` と `class MidiToKeyboard` が同名で存在するため、型解決で衝突する場合がある。
- `MidiToKeyboard.Domain` や `MidiToKeyboard.Application` を参照する箇所で解決できない場合、`global::MidiToKeyboard...` が必要か確認する。
- 不要な `global::` は増やさず、名前衝突を避けるために必要な箇所だけで使う。

### MIDIデバイスが見つからない場合

- Windows のデバイス マネージャーや MIDI 対応アプリで、対象デバイスが認識されているか確認する。
- 他のアプリが MIDI デバイスを占有していないか確認する。
- USB MIDI デバイスの場合、接続し直してからアプリを再起動する。
- Visual Studio から再実行し、デバイス一覧に表示されるか確認する。

### mappings.json が見つからない場合

- プロジェクト直下に `mappings.json` が存在するか確認する。
- Visual Studio のプロパティで、出力ディレクトリへコピーされる設定になっているか確認する。
- `bin\Debug` または `bin\Release` に `mappings.json` がコピーされているか確認する。
- 見つからない場合の既定マッピング使用表示が、既存挙動どおりか確認する。

### SendInput が失敗する場合

- 既存 Main 経路では、`WindowsKeyOutput.WarningOccurred` の購読により console に `GetLastWin32Error` 付きのエラー表示が出ているか確認する。
- `WarningOccurred` が未購読の経路では、失敗通知は表示されないが例外にならないことを確認する。
- 対象アプリが管理者権限で動いている場合、権限差により入力が届かない可能性を確認する。
- 入力先ウィンドウにフォーカスがあるか確認する。
- セキュリティソフトや OS の入力制限により SendInput がブロックされていないか確認する。
- `VirtualKey` と `Scancode` の両方の入力送信モードで挙動を比較する。
