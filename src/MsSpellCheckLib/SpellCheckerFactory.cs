using System.Runtime.InteropServices;

namespace MsSpellCheckLib
{
    [CoClass(typeof(SpellCheckerFactoryClass))]
    [Guid("8E018A9D-2415-4677-BF08-794EA61F94BB")]
    [ComImport]
    public interface SpellCheckerFactory : ISpellCheckerFactory
    {
    }
}
