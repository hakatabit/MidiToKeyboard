using System;
using System.IO;
using MidiToKeyboard.Application;
using MidiToKeyboard.Infrastructure;
using MidiToKeyboard.Ui;

namespace MidiToKeyboard
{
    /// <summary>
    /// MIDI 入力をキーボード入力に変換して送信するコンソールアプリケーションのエントリポイント
    /// </summary>
    internal static class Program
    {
        private const string MappingFileName = "mappings.json";

        private static void Main(string[] args)
        {
            try
            {
                string mappingsPath = Path.Combine(
                    AppDomain.CurrentDomain.BaseDirectory,
                    MappingFileName);

                using (var midiInput = new DryWetMidiInput())
                {
                    var keyOutput = new WindowsKeyOutput();
                    var profileRepository = new JsonProfileRepository(mappingsPath);
                    var application = new MidiToKeyboardApplication(
                        midiInput,
                        keyOutput,
                        profileRepository);
                    var consoleUi = new ConsoleUi(
                        application,
                        midiInput,
                        profileRepository);

                    keyOutput.WarningOccurred += consoleUi.HandleWarning;
                    application.MidiInputActivityOccurred += consoleUi.HandleMidiInputActivity;

                    try
                    {
                        consoleUi.Run();
                    }
                    finally
                    {
                        application.MidiInputActivityOccurred -= consoleUi.HandleMidiInputActivity;
                        keyOutput.WarningOccurred -= consoleUi.HandleWarning;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("エラーが発生しました: {0}", ex.Message);
            }
        }
    }
}
