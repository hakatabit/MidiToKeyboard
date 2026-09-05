# MidiToKeyboard 設計方針（Refactor 版）

## 目的
- コンソール UI から **タスクトレイ常駐 UI** に移行しても、内部ロジックをほぼ変更せずに差し替えできる構造にする。
- MIDI 受信 → 変換（マッピング/状態）→ キー送出 を、UI / ライブラリ / OS 依存から分離する。
- 小さな変更単位で安全にリファクタできるよう、依存方向と責務境界を固定する。

## 非目標（この段階ではやらない）
- SysEx 対応
- 特定ハードウェア（音源・MIDI機器）に依存した制御
- 複数デバイス / 複数トラックの高度なルーティング（将来の拡張で対応）
- UI デザインの刷新（まずは差し替え可能な構造を作る）

## アーキテクチャ（依存方向）
依存方向は次のとおりとする。

- Ui → Application
- Application → Domain
- Infrastructure → Application の interface / Domain
- Domain は他レイヤに依存しない
- Application は Infrastructure の具象クラスに直接依存しない
- Program は Composition Root として具象クラスの生成と配線を担当する

- **Composition Root（Program）**
  - 具象クラスの生成を担当する
  - イベント購読の配線を担当する
  - `ConsoleUi.Run()` を呼び出す
  - Console 表示ロジックは原則として持たない

- **UI**
  - コンソール / タスクトレイ / 将来の設定画面など
  - アプリの開始 / 停止、プロファイル選択、ログ表示を担当
  - ロジックを持たない（Application へ命令するだけ）

- **Application**
  - ユースケース（Start / Stop、プロファイル適用、状態通知）の司令塔
  - Domain を呼び、Infrastructure を利用する
  - UI から見える API を提供（UI を差し替え可能にする）

- **Domain**
  - MIDI イベントの解釈、マッピング、押下状態などの **純粋ロジック**
  - OS / ライブラリ / ファイル I/O を知らない
  - テストしやすい（入力 → 出力が決まる）

- **Infrastructure**
  - DryWetMIDI による MIDI 入力
  - Win32 SendInput によるキー送出
  - JSON の読み込みなど
  - 外部依存はここに隔離する

## モジュール / 責務（主要コンポーネント）

### Composition Root
- `Program`
  - `DryWetMidiInput`、`WindowsKeyOutput`、`JsonProfileRepository`、`MidiToKeyboardApplication`、`ConsoleUi` を生成する
  - 通知イベントを `ConsoleUi` の表示ハンドラへ接続する
  - `ConsoleUi.Run()` を呼ぶ

### UI
- `ConsoleUi`
  - Console 入出力を担当する
  - MIDI デバイス選択、入力送信モード選択、プロファイル選択を担当する
  - MIDI ON / OFF と `WindowsKeyOutput` の警告を表示する
  - `MidiToKeyboardApplication` の Start / Stop を呼ぶ

### Application
- `MidiToKeyboardApplication`
  - `Start(deviceId, profileName, inputMode)` で指定された MIDI デバイス、Profile、入力送信モードを使用して処理を開始し、`Stop()` で停止する
  - `SetProfile(profileName)` で指定された Profile を読み込み、MIDI イベントの変換に適用する
  - MIDI入力イベントを受け、Domain の変換結果を KeyOutput に流す
  - `MidiInputActivityOccurred` で MIDI 入力活動を UI に通知する
- `IMidiInput`
  - MIDI デバイス列挙、受信開始 / 停止、MIDI メッセージ通知を定義する
- `IKeyOutput`
  - `Send(MidiToKeyboard.Domain.KeyAction action)` を定義する
- `IInputModeKeyOutput`
  - `SetInputMode(InputMode inputMode)` を定義する
- `IProfileRepository`
  - Profile の一覧取得と読み込みを定義する
- `InputMode`
  - VirtualKey / Scancode の送信方式を表す
- `MidiDeviceInfo`
  - MIDI デバイスの ID と表示名を表す
- `MidiInputActivity` / `MidiInputActivityType`
  - UI に通知する MIDI 入力活動を表す

### Domain
- `MidiEvent` / `MidiEventType`
  - ライブラリ非依存の MIDI 表現（NoteOn / NoteOff / CC など）
- `KeyAction` / `KeyActionType`
  - ライブラリ非依存のキー操作表現（KeyDown / KeyUp / KeyPress / UnicodeText）
- `Profile`
  - マッピングと関連設定の集合
- `KeyPressState`
  - 重複 KeyUp 防止、同時押しの参照カウント管理
- `MidiTranslator`
  - `MidiEvent` を解釈し、`KeyAction` を生成する

### Infrastructure
- `DryWetMidiInput : IMidiInput`
  - DryWetMIDI をここに閉じ込める
- `WindowsKeyOutput : IKeyOutput, IInputModeKeyOutput`
  - SendInput / PInvoke をここに閉じ込める
  - SendInput 失敗時は `WarningOccurred` を発火し、Console へ直接出力しない
- `JsonProfileRepository : IProfileRepository`
  - mappings.json の読み込みをここに閉じ込める
- `NativeMethods`
  - SendInput などの Win32 API 宣言をここに閉じ込める

## InputMode と KeyActionType の役割分担
- `KeyActionType` は「何を押すか」を表す
- `InputMode` は「どう送るか」を表す
- VirtualKey / Scancode の切り替えは `InputMode` で扱う
- `KeyActionType` に送信方式を混ぜない

## 通知設計
- `MidiToKeyboardApplication` は `MidiInputActivityOccurred` で MIDI 入力活動を通知する
- `WindowsKeyOutput` は `WarningOccurred` でキー送出警告を通知する
- `Program` が各イベントを `ConsoleUi` のハンドラへ接続する
- `ConsoleUi` が通知内容を Console に表示する
- Application / Infrastructure は `Console.WriteLine` を直接呼ばない

## mappings 未登録時の方針
- 選択可能な Profile がない場合はメッセージを表示して安全に終了する
- mappings.json が存在しない、または読み込めない場合はエラーメッセージを表示して終了する
- 代替マッピングを自動生成しない

## リファクタのガードレール
- Domain から Infrastructure を参照しない
- UI は Application 経由でのみ処理を行う
- 挙動を変えない（NoteOn velocity=0 → NoteOff 等）

## 互換性メモ
- .NET Framework 4.7.2（旧形式 csproj）
- ビルド / 実行は Windows（Visual Studio）前提
- C# 10 以上の構文は使用しない
- ファイルスコープ namespace は使用しない

## 今後の予定
- コンソール UI からタスクトレイ UI への移行
- UI 差し替え時も Domain / Infrastructure の変更を小さく保つ
- 必要に応じた自動テストの追加を検討する
