namespace MidiToKeyboard.Application
{
    public interface IKeyOutput
    {
        void Send(global::MidiToKeyboard.Domain.KeyAction action);
    }
}
