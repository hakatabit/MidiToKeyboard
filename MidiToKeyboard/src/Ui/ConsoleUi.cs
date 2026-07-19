using System;

namespace MidiToKeyboard.Ui
{
    public sealed class ConsoleUi
    {
        private readonly MidiToKeyboard.Application.MidiToKeyboardApplication _application;

        public ConsoleUi(MidiToKeyboard.Application.MidiToKeyboardApplication application)
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
