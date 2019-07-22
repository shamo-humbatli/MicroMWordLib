using System.Reflection;

namespace MicroMWordLib.WordOperations
{
    public abstract class WParameters
    {
        private static System.Reflection.Missing ref_Missing;

        public static Missing Missing { get => ref_Missing; }
    }
}
