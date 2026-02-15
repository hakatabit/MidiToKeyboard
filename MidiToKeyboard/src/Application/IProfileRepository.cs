using System.Collections.Generic;

namespace MidiToKeyboard.Application
{
    public interface IProfileRepository
    {
        global::MidiToKeyboard.Domain.Profile Load(string name);
        IEnumerable<string> ListNames();
    }
}
