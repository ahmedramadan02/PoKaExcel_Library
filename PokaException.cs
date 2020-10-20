using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace pokaLibrary
{
    /*
    Defining Exception Classes 
     * Programs can throw a predefined exception class in the System namespace 
     * (except where previously noted), or create their own exception classes by deriving from Exception. 
     * The derived classes should define at least four constructors: one default constructor, 
     * one that sets the message property, and one that sets both the Message and InnerException properties. 
     * The fourth constructor is used to serialize the exception. New exception classes should be serializable
    */
    public class PokaException : Exception
    {
        public PokaException(string message) : base(message) { }
    }
}
