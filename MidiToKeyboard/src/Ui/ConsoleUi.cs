using System;

namespace MidiToKeyboard.Ui
{
    public sealed class ConsoleUi
    {
        private readonly global::MidiToKeyboard.Application.MidiToKeyboardApplication _application;

        public ConsoleUi(global::MidiToKeyboard.Application.MidiToKeyboardApplication application)
        {
            if (application == null)
                throw new ArgumentNullException(nameof(application));

            _application = application;
        }

        public void Run()
        {
        }
    }
}
