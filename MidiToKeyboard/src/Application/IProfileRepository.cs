using System.Collections.Generic;

namespace MidiToKeyboard.Application
{
    public interface IProfileRepository
    {
        MidiToKeyboard.Domain.Profile Load(string name);
        IEnumerable<string> ListNames();
    }
}
