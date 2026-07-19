using System;
using System.Collections.Generic;

namespace MidiToKeyboard.Application
{
    public interface IMidiInput
    {
        event Action<MidiToKeyboard.Domain.MidiEvent> MessageReceived;
        IReadOnlyList<MidiDeviceInfo> EnumerateDevices();
        void Start(string deviceId);
        void Stop();
    }
}
