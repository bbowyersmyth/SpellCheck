// <copyright file="UserDictionariesRegistrar.cs" company="Bruce Bowyer-Smyth">
//     Copyright © Bruce Bowyer-Smyth
// </copyright>

namespace PlatformSpellCheck
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// Manages the registration of user dictionaries
    /// Implements the <see cref="System.IDisposable" />
    /// </summary>
    /// <seealso cref="System.IDisposable" />
    public class UserDictionariesRegistrar : IDisposable
    {
        /// <summary>
        /// The registrar
        /// </summary>
        private MsSpellCheckLib.IUserDictionariesRegistrar registrar;

        /// <summary>
        /// Maintains user dictionaries
        /// </summary>
        public UserDictionariesRegistrar() => this.registrar =
            (MsSpellCheckLib.IUserDictionariesRegistrar)new MsSpellCheckLib.SpellCheckerFactory();

        /// <summary>
        /// Registers a file to be used as a user dictionary for the current user, until unregistered
        /// </summary>
        /// <param name="dictionaryPath">The path of the dictionary file to be registered</param>
        /// <param name="languageTag">The language for which this dictionary should be used. If left empty, it will be used for any language</param>
        public void RegisterUserDictionary(string dictionaryPath, string languageTag) =>
            this.registrar.RegisterUserDictionary(dictionaryPath, languageTag);

        /// <summary>
        /// Unregisters a previously registered user dictionary. The dictionary will no longer be used by the spell checking functionality.
        /// </summary>
        /// <param name="dictionaryPath">The path of the dictionary file to be unregistered</param>
        /// <param name="languageTag">
        /// The language for which this dictionary was used. It must match the language passed to <see cref="RegisterUserDictionary(string, string)" />
        /// </param>
        public void UnregisterUserDictionary(string dictionaryPath, string languageTag) =>
            this.registrar.UnregisterUserDictionary(dictionaryPath, languageTag);

        /// <summary>
        /// Disposes resources used by UserDictionariesRegistrar
        /// </summary>
        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Disposes resources used by UserDictionariesRegistrar
        /// </summary>
        ~UserDictionariesRegistrar()
        {
            this.Dispose(false);
        }

        /// <summary>
        /// Releases unmanaged and - optionally - managed resources.
        /// </summary>
        /// <param name="disposing"><c>true</c> to release both managed and unmanaged resources; <c>false</c> to release only unmanaged resources.</param>
        protected virtual void Dispose(bool disposing)
        {
            if (this.registrar != null)
            {
                Marshal.ReleaseComObject(this.registrar);
                this.registrar = null;
            }
        }
    }
}
