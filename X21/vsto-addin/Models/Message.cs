using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Text.Json.Serialization;
using System.Collections.Generic;

namespace X21.Models
{
    /// <summary>
    /// Represents a message in conversations (chat, LLM, etc.)
    /// </summary>
    public class Message : INotifyPropertyChanged
    {
        private string _role;
        private string _content;
        private DateTime _timestamp;
        private string _traceId;

        public event PropertyChangedEventHandler PropertyChanged;

        [JsonPropertyName("role")]
        public string Role
        {
            get => _role;
            set => SetProperty(ref _role, value);
        }

        [JsonPropertyName("content")]
        public string Content
        {
            get => _content;
            set => SetProperty(ref _content, value);
        }

        [JsonPropertyName("timestamp")]
        public DateTime Timestamp
        {
            get => _timestamp;
            set => SetProperty(ref _timestamp, value);
        }

        [JsonPropertyName("traceId")]
        public string TraceId
        {
            get => _traceId;
            set => SetProperty(ref _traceId, value);
        }

        public Message(string role, string content, DateTime timestamp, string traceId)
        {
            Role = role;
            Content = content;
            Timestamp = timestamp;
            TraceId = traceId;
        }

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        protected virtual bool SetProperty<T>(ref T field, T value, [CallerMemberName] string propertyName = null)
        {
            if (Equals(field, value))
                return false;

            field = value;
            OnPropertyChanged(propertyName);
            return true;
        }
    }
}
