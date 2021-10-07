// From Source Code Analysis with Roslyn -

using System;

public class PC_ValueTypesOverrideEqualsAndGetHashCode
{
    struct Vector : IEquatable<Vector>
    {
        public int X { get; set; }
        public int Y { get; set; }
        public int Z { get; set; }
        public int Magnitude { get; set; }
        public override bool Equals(object obj)
        {
            if (obj == null)
            {
                return false;
            }
            if (obj.GetType() != this.GetType())
            {
                return false;
            }
            return this.Equals((Vector)obj);
        }
        public bool Equals(Vector other)
        {
            return this.X == other.X
            && this.Y == other.Y
            && this.Z == other.Z
            && this.Magnitude == other.Magnitude;
        }
        //Deliberately commented to make this struct a “defaulter”
        //public override int GetHashCode()
        //{
        // return X ^ Y ^ Z ^ Magnitude;
        //}
    }
}
