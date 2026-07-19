# DESIGN.md
MidiToKeyboard 設計方針（Refactor 版）

## 目的
- コンソール UI から **タスクトレイ常駐 UI** に移行しても、内部ロジックをほぼ変更せずに差し替えできる構造にする。
- MIDI 受信 → 変換（マッピング/状態）→ キー送出 を、UI / ライブラリ / OS 依存から分離する。
- 小さな変更単位で安全にリファクタできるよう、依存方向と責務境界を固定する。

## 非目標（この段階ではやらない）
- SysEx 対応
- 特定ハードウェア（音源・MIDI機器）に依存した制御
- 複数デバイス / 複数トラックの高度なルーティング（将来の拡張で対応）
- UI デザインの刷新（まずは差し替え可能な構造を作る）

---

## アーキテクチャ（依存方向）
依存は **上から下へだけ**。

UI → Application → Domain → Infrastructure

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
  - JSON の読み書き、ログ実装など
  - 外部依存はここに隔離する

---

## モジュール / 責務（主要コンポーネント）

### UI
- `ConsoleUi`
  - デバイス一覧表示、選択、開始 / 停止、プロファイル選択
  - `MidiToKeyboard`（Application）の API のみを呼ぶ
- `TrayUi`（将来）
  - タスクトレイ常駐、メニュー（Start / Stop / Profile / Exit）
  - `MidiToKeyboard` の API のみを呼ぶ

### Application
- `MidiToKeyboard`
  - `Start(deviceId, profileName, inputMode)` / `Stop()`
  - `SetProfile(profileName)`（将来のトレイ切替用）
  - MIDI入力イベントを受け、Domain の変換結果を KeyOutput に流す
  - 状態通知（イベント or コールバック）を UI に返す

### Domain
- `MidiEvent`
  - ライブラリ非依存の MIDI 表現（NoteOn / NoteOff / CC など）
- `KeyAction`
  - ライブラリ非依存のキー送出表現（KeyDown / KeyUp / Unicode / Scancode / VK）
- `Profile`
  - マッピングと関連設定の集合
- `KeyPressState`
  - 重複 KeyUp 防止、同時押しの参照カウント管理
- `MidiTranslator`
  - `MidiEvent` を解釈し、`KeyAction` を生成する

### Infrastructure
- `IMidiInput`
  - `event Action<MidiToKeyboard.Domain.MidiEvent> MessageReceived`
- `DryWetMidiInput : IMidiInput`
  - DryWetMIDI をここに閉じ込める

- `IKeyOutput`
  - `Send(MidiToKeyboard.Domain.KeyAction action)`
- `WindowsKeyOutput : IKeyOutput`
  - SendInput / PInvoke をここに閉じ込める

- `IProfileRepository`
- `JsonProfileRepository : IProfileRepository`

---

## リファクタのガードレール
- Domain から Infrastructure を参照しない
- UI は Application 経由でのみ処理を行う
- 挙動を変えない（NoteOn velocity=0 → NoteOff 等）

---

## 互換性メモ
- .NET Framework 4.7.2（旧形式 csproj）
- ビルド / 実行は Windows（Visual Studio）前提
