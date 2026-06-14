using Melanchall.DryWetMidi.Core;
using Melanchall.DryWetMidi.Multimedia;
using MidiToKeyboard.Application;
using System;
using System.Collections.Generic;
using System.Linq;
using DomainMidiEvent = MidiToKeyboard.Domain.MidiEvent;

namespace MidiToKeyboard.Infrastructure
{
    public sealed class DryWetMidiInput : IMidiInput, IDisposable
    {
        private InputDevice _inputDevice;

        public event Action<DomainMidiEvent> MessageReceived;

        public IReadOnlyList<MidiDeviceInfo> EnumerateDevices()
        {
            var devices = InputDevice.GetAll().ToList();

            try
            {
                return devices
                    .Select(device => new MidiDeviceInfo(device.Name, device.Name))
                    .ToList();
            }
            finally
            {
                foreach (var device in devices)
                {
                    device.Dispose();
                }
            }
        }

        public void Start(string deviceId)
        {
            Stop();

            var inputDevice = InputDevice.GetByName(deviceId);

            try
            {
                inputDevice.EventReceived += OnEventReceived;
                inputDevice.StartEventsListening();
                _inputDevice = inputDevice;
            }
            catch
            {
                inputDevice.EventReceived -= OnEventReceived;
                inputDevice.Dispose();
                throw;
            }
        }

        public void Stop()
        {
            var inputDevice = _inputDevice;
            if (inputDevice == null)
            {
                return;
            }

            _inputDevice = null;
            inputDevice.EventReceived -= OnEventReceived;

            try
            {
                inputDevice.StopEventsListening();
            }
            finally
            {
                inputDevice.Dispose();
            }
        }

        public void Dispose()
        {
            Stop();
        }

        private void OnEventReceived(object sender, MidiEventReceivedEventArgs e)
        {
            DomainMidiEvent midiEvent = ConvertEvent(e.Event);
            if (midiEvent != null)
            {
                MessageReceived?.Invoke(midiEvent);
            }
        }

        private static DomainMidiEvent ConvertEvent(Melanchall.DryWetMidi.Core.MidiEvent midiEvent)
        {
            var noteOnEvent = midiEvent as NoteOnEvent;
            if (noteOnEvent != null)
            {
                int channel = noteOnEvent.Channel;
                int noteNumber = noteOnEvent.NoteNumber;
                int velocity = noteOnEvent.Velocity;

                return velocity == 0
                    ? DomainMidiEvent.NoteOff(channel, noteNumber)
                    : DomainMidiEvent.NoteOn(channel, noteNumber, velocity);
            }

            var noteOffEvent = midiEvent as NoteOffEvent;
            if (noteOffEvent != null)
            {
                return DomainMidiEvent.NoteOff(
                    noteOffEvent.Channel,
                    noteOffEvent.NoteNumber);
            }

            var controlChangeEvent = midiEvent as ControlChangeEvent;
            if (controlChangeEvent != null)
            {
                return DomainMidiEvent.ControlChange(
                    controlChangeEvent.Channel,
                    controlChangeEvent.ControlNumber,
                    controlChangeEvent.ControlValue);
            }

            return null;
        }
    }
}
