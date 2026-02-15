namespace MidiToKeyboard.Application
{
    public sealed class MidiDeviceInfo
    {
        public string Id { get; }
        public string Name { get; }

        public MidiDeviceInfo(string id, string name)
        {
            Id = id;
            Name = name;
        }
    }
}
