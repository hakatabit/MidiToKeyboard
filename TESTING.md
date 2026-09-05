# MidiToKeyboard テストガイド

## 目的

この文書は、現在のUi / Application / Domain / Infrastructure構成について、WindowsのVisual Studio上で確認するビルド観点と手動動作確認観点をまとめたものです。

現時点では自動テストが整備されていないため、リファクタや機能変更後は本書の観点を使用して既存挙動を確認します。

## 前提環境

- Windows
- Visual Studio
- .NET Framework 4.7.2
- 旧形式csproj
- Windowsから認識できるMIDI入力デバイス
- キー入力を確認できるテキストエディタまたは対象アプリケーション
- 必要なNuGetパッケージが復元済みであること

WSL上でのビルドや実行確認は前提にしません。VirtualKey / Scancodeと`SendInput`の確認はWindows上で行います。

## ビルド確認

1. Visual Studioで`MidiToKeyboard.sln`を開く。
2. NuGetパッケージが復元されていることを確認する。
3. Debug / Any CPU構成でビルドする。
4. Release / Any CPU構成でビルドする。
5. 対象フレームワークが.NET Framework 4.7.2であることを確認する。
6. ビルドエラーと新しい警告がないことを確認する。
7. `MidiToKeyboard.csproj`が旧形式csprojとして読み込めることを確認する。
8. `Compile Include`にProgramと必要な`src`配下のファイルが含まれていることを確認する。
9. 存在しないファイルや削除済みファイルへの`Compile Include`が残っていないことを確認する。
10. `mappings.json`の`CopyToOutputDirectory`が有効であることを確認する。
11. `bin\Debug`と`bin\Release`へ`mappings.json`がコピーされることを確認する。

## 起動・Composition Root確認

1. `Program.Main`が次の具象クラスを生成していることを確認する。
   - `DryWetMidiInput`
   - `WindowsKeyOutput`
   - `JsonProfileRepository`
   - `MidiToKeyboardApplication`
   - `ConsoleUi`
2. `WindowsKeyOutput.WarningOccurred`が`ConsoleUi.HandleWarning`へ接続されていることを確認する。
3. `MidiToKeyboardApplication.MidiInputActivityOccurred`が`ConsoleUi.HandleMidiInputActivity`へ接続されていることを確認する。
4. `Program.Main`が`ConsoleUi.Run()`を呼び出していることを確認する。
5. Visual StudioからDebug実行し、起動直後に想定外の例外で終了しないことを確認する。
6. MIDI入力デバイスがない場合は、デバイスが見つからない旨を表示して安全に終了することを確認する。

## 操作フロー確認

表示と入力の順序が次のとおりであることを確認します。

1. MIDI入力デバイス一覧が表示される。
2. MIDI入力デバイスを番号で選択できる。
3. 入力送信モードの選択肢が表示される。
   - `[1]` VirtualKey
   - `[2]` Scancode
4. `2`を入力すると`InputMode.Scancode`が選択される。
5. `2`以外、空入力、不正入力では既定の`InputMode.VirtualKey`が選択される。
6. プロファイル一覧が表示される。
7. プロファイルを番号で選択できる。
8. 無効なプロファイル番号では先頭のプロファイルが選択される。
9. 選択したデバイス名と入力モードが開始メッセージに表示される。
10. MIDI入力待ち状態になる。
11. コンソールで任意のキーを押すと終了する。
12. 終了時に`MidiToKeyboardApplication.Stop()`が呼ばれる。

無効なMIDIデバイス番号を入力した場合は、無効な選択である旨を表示し、MIDI受信を開始せず安全に終了することを確認します。

## InputMode反映確認

1. `ConsoleUi`が選択した`InputMode`を3引数の`MidiToKeyboardApplication.Start()`へ渡すことを確認する。
2. Applicationが`IInputModeKeyOutput.SetInputMode()`を呼ぶことを確認する。
3. `WindowsKeyOutput`が選択された`InputMode`を保持することを確認する。
4. VirtualKeyでは`SendKeyInput`経路が使用されることを確認する。
5. Scancodeでは`SendScancodeKey`経路が使用されることを確認する。
6. Domainの`KeyActionType`にVirtualKey / Scancodeが含まれていないことを確認する。
7. 入力モードの変更によってDomainの変換判断が変わらないことを確認する。

## `mappings.json` / Profile確認

### 正常系

1. `mappings.json`が実行ファイルと同じディレクトリに存在することを確認する。
2. JSONの`sets`からプロファイル名を読み込めることを確認する。
3. 複数プロファイルを定義した場合、すべて一覧表示されることを確認する。
4. 選択したプロファイル名が表示されることを確認する。
5. 選択したプロファイルのノート番号とキー文字の対応が適用されることを確認する。
6. `mappings.json`のJSON構造がREADME記載の形式から変わっていないことを確認する。

### 未登録・異常系

1. `sets`が空の場合、プロファイルがない旨のメッセージを表示して終了することを確認する。
2. `mappings.json`が存在しない場合、エラーメッセージを表示して終了することを確認する。
3. JSONを解析できない場合、エラーメッセージを表示して終了することを確認する。
4. いずれの場合も代替マッピングや既定プロファイルが自動生成されないことを確認する。
5. 終了時に`Stop()`とMIDI入力の破棄が実行されることを確認する。
6. 選択したプロファイルの`mappings`が空の場合、不要なキー送出や想定外の例外が発生しないことを確認する。
7. 数値として解釈できないノート番号と空のキー文字列が読み飛ばされることを確認する。
8. `JsonProfileRepository.Load()`へ存在しない名前を直接指定した場合、プロファイル未検出の例外になることを確認する。

## MIDI入力確認

1. Windowsが対象MIDIデバイスを認識していることを確認する。
2. 他のアプリケーションがデバイスを占有していないことを確認する。
3. アプリケーションからデバイスを選択し、受信待ち状態にする。
4. NoteOnを送信し、MIDI ON活動がConsoleUiに表示されることを確認する。
5. NoteOffを送信し、MIDI OFF活動がConsoleUiに表示されることを確認する。
6. 表示されるNoteNumberが入力したノート番号と一致することを確認する。
7. 表示されるKeyCharが選択したプロファイルの対応と一致することを確認する。
8. Velocity 0のNoteOnがNoteOff相当に変換され、MIDI OFFとして扱われることを確認する。
9. ControlChangeを送信してもキー操作が生成されず、想定外の例外で終了しないことを確認する。
10. 未登録ノートを送信しても不要なキー操作が生成されないことを確認する。

## キー送出確認

### 基本動作

1. キー入力を受け取る対象アプリケーションへフォーカスを移す。
2. マッピング済みノートのNoteOnでKeyDown相当が送出されることを確認する。
3. 同じノートのNoteOffでKeyUp相当が送出されることを確認する。
4. Velocity 0のNoteOnでもKeyUp相当が送出されることを確認する。
5. VirtualKeyモードで対象アプリケーションにキー入力が届くことを確認する。
6. Scancodeモードで対象アプリケーションにキー入力が届くことを確認する。
7. `KeyPress`を使用する経路を追加・変更した場合は、KeyDownからKeyUpの順に送出されることを確認する。

現在の`MidiTranslator`はNoteOnから`KeyDown`、NoteOffから`KeyUp`を生成します。`KeyPress`は通常のMIDI変換経路では使用していません。

### 状態管理

1. 同じノートのNoteOnを複数回受信しても、不要なKeyDownが連続送出されないことを確認する。
2. 同じノートのNoteOffを複数回受信しても、不要なKeyUpが連続送出されないことを確認する。
3. 同じキーへ複数のノートを割り当てる。
4. 1つ目のノートでKeyDownが送出されることを確認する。
5. 2つ目のノートで重複KeyDownが送出されないことを確認する。
6. 片方のノートだけを離した時点ではKeyUpが送出されないことを確認する。
7. 両方のノートを離した時点でKeyUpが送出されることを確認する。

## 警告表示確認

1. `WindowsKeyOutput`に`Console.WriteLine` / `Console.Write`がないことを確認する。
2. `SendInput`の戻り値が0の場合に`WarningOccurred`が発火することを確認する。
3. 警告メッセージに`GetLastWin32Error`の値が含まれることを確認する。
4. `Program`が`WarningOccurred`を`ConsoleUi.HandleWarning`へ接続していることを確認する。
5. `ConsoleUi.HandleWarning`が受け取ったメッセージを表示することを確認する。
6. イベントが未購読でも、警告通知処理自体が例外にならないことを確認する。

可能であれば、`SendInput`が拒否される権限差などを再現し、実行時にも警告がConsoleUi経由で表示されることを確認します。再現できない場合は、イベント発火と購読経路をコードレビューで確認します。

## 終了・例外時の確認

1. 通常終了時に`ConsoleUi.Run()`の`finally`から`Stop()`が呼ばれることを確認する。
2. デバイス列挙、プロファイル列挙、Start、終了待ちのいずれかで例外が発生しても、`Stop()`が呼ばれることを確認する。
3. ApplicationがMIDI入力の`MessageReceived`購読を解除することを確認する。
4. `DryWetMidiInput.Stop()`が受信を停止し、入力デバイスを破棄することを確認する。
5. `Program`がMIDI活動・警告イベントの購読を解除することを確認する。
6. `DryWetMidiInput.Dispose()`が呼ばれることを確認する。
7. Mainまで伝播した例外は、メッセージを表示して終了することを確認する。

## レイヤ確認

1. DomainがConsole、DryWetMIDI、Windows API、JSON、Ui、Application、Infrastructureに依存していないことを確認する。
2. Applicationに`Console.WriteLine` / `Console.ReadLine` / `Console.ReadKey`がないことを確認する。
3. ApplicationがInfrastructureの具象クラスを`new`していないことを確認する。
4. Applicationが`IMidiInput` / `IKeyOutput` / `IProfileRepository`などの抽象を利用していることを確認する。
5. InfrastructureがUiに依存していないことを確認する。
6. `WindowsKeyOutput`がWin32依存を閉じ込めていることを確認する。
7. `DryWetMidiInput`がDryWetMIDI依存を閉じ込めていることを確認する。
8. `JsonProfileRepository`がJSONとファイルI/O依存を閉じ込めていることを確認する。
9. `Program`がComposition Rootとして具象クラス生成とイベント配線を担当していることを確認する。
10. `ConsoleUi`がConsole入出力を担当していることを確認する。
11. `ConsoleUi`が`KeyPressState`や`MidiTranslator`の内部状態を直接操作していないことを確認する。
12. 不要な`global::`修飾がないことを確認する。

## トラブルシュート

### MIDIデバイスが表示されない

- Windowsのデバイスマネージャーや他のMIDI対応アプリで認識されているか確認する。
- 他のアプリケーションがデバイスを占有していないか確認する。
- USB MIDIデバイスを接続し直し、アプリケーションを再起動する。

### `mappings.json`を読み込めない

- プロジェクト直下に`mappings.json`が存在するか確認する。
- csprojの`CopyToOutputDirectory`設定を確認する。
- `bin\Debug`または`bin\Release`にファイルがコピーされているか確認する。
- JSON構文と`sets` / `name` / `mappings`の構造を確認する。
- 現在の実装は代替マッピングを生成しないため、エラーを修正してから再起動する。

### キー入力が届かない

- 入力先ウィンドウにフォーカスがあるか確認する。
- VirtualKeyとScancodeの両方で比較する。
- 対象アプリケーションとMidiToKeyboardの実行権限に差がないか確認する。
- OSやセキュリティソフトが`SendInput`を制限していないか確認する。
- Consoleに`GetLastWin32Error`付きの警告が表示されていないか確認する。

## 既知の未対応事項

- タスクトレイUIは未対応で、現在はコンソールUIのみです。
- 自動テストは未整備です。
- MIDI入力とWindows `SendInput`を使用するためWindows専用です。
- WSL上でのビルド・実行確認は前提外です。
- SysExには対応していません。
- 複数デバイス・複数トラックの高度なルーティングには対応していません。
