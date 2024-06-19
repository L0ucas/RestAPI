using System;
using System.Runtime.Serialization;

namespace TMHelper.Common
{
    [Serializable]
    public class BusinessRuleException : Exception
    {
        public BusinessRuleException() { }

        public BusinessRuleException(string message) : base(message) { }

        public BusinessRuleException(string message, Exception innerException) : base(message, innerException) { }

        protected BusinessRuleException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}