# SpellCheck
Managed wrapper for the Microsoft Spell Checking API available in Windows 8 and Windows Server 2012 and later.

https://www.nuget.org/packages/PlatformSpellCheck

    var spelling = new SpellChecker();
    
    foreach (var mistake in spelling.Check("speelling"))
    {
      Console.WriteLine("Start: {0} Length: {1}", mistake.StartIndex, mistake.Length);
    }
    
    foreach (var word in spelling.Suggestions("speelling"))
    {
      Console.WriteLine(word);
    }
  
    spelling.Dispose();
