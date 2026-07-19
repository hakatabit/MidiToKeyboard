namespace MidiToKeyboard.Application
{
    public interface IKeyOutput
    {
        void Send(MidiToKeyboard.Domain.KeyAction action);
    }
}
