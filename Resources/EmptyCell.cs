public struct EmptyCell
{
    public static dynamic operator +(dynamic a, EmptyCell b)
    {
        return a;
    }

    public static dynamic operator +(EmptyCell a, dynamic b)
    {
        return b;
    }

    public static dynamic operator +(EmptyCell a, EmptyCell b)
    {
        return a;
    }

    public static dynamic operator -(dynamic a, EmptyCell b)
    {
        return a;
    }

    public static dynamic operator -(EmptyCell a, dynamic b)
    {
        return 0 - b;
    }

    public static dynamic operator -(EmptyCell a, EmptyCell b)
    {
        return 0;
    }

    public static dynamic operator *(dynamic a, EmptyCell b)
    {
        return 0;
    }

    public static dynamic operator *(EmptyCell a, dynamic b)
    {
        return 0;
    }

    public static dynamic operator *(EmptyCell a, EmptyCell b)
    {
        return 0;
    }

    public static dynamic operator /(dynamic a, EmptyCell b)
    {
        throw new System.DivideByZeroException();
    }

    public static dynamic operator /(EmptyCell a, dynamic b)
    {
        return 0;
    }

    public static dynamic operator /(EmptyCell a, EmptyCell b)
    {
        throw new System.DivideByZeroException();
    }

    public static dynamic operator %(dynamic a, EmptyCell b)
    {
        throw new System.DivideByZeroException();
    }

    public static dynamic operator %(EmptyCell a, dynamic b)
    {
        return 0;
    }

    public static dynamic operator %(EmptyCell a, EmptyCell b)
    {
        throw new System.DivideByZeroException();
    }

    public static bool IsNumeric(object o) => o is byte || o is sbyte || o is ushort || o is uint || o is ulong ||
                                              o is short || o is int || o is long || o is float || o is double ||
                                              o is decimal;

    public static bool operator ==(dynamic a, EmptyCell empty)
    {
        if (IsNumeric(a)) return a == 0;
        switch (a)
        {
            case string s:
                return s == "";
            case bool b:
                return b == false;
            default:
                return false;
        }
    }

    public static bool operator ==(EmptyCell a, dynamic b)
    {
        return b == a;
    }

    public static bool operator ==(EmptyCell a, EmptyCell b)
    {
        return true;
    }

    public static bool operator !=(dynamic a, EmptyCell b)
    {
        return !(a == b);
    }

    public static bool operator !=(EmptyCell a, dynamic b)
    {
        return !(b == a);
    }

    public static bool operator !=(EmptyCell a, EmptyCell b)
    {
        return false;
    }

    public static bool operator <(dynamic a, EmptyCell b)
    {
        return a < 0;
    }

    public static bool operator <(EmptyCell a, dynamic b)
    {
        return 0 < b;
    }

    public static bool operator <(EmptyCell a, EmptyCell b)
    {
        return false;
    }

    public static bool operator >(dynamic a, EmptyCell b)
    {
        return a > 0;
    }

    public static bool operator >(EmptyCell a, dynamic b)
    {
        return 0 > b;
    }

    public static bool operator >(EmptyCell a, EmptyCell b)
    {
        return false;
    }

    public static bool operator <=(dynamic a, EmptyCell b)
    {
        return a <= 0;
    }

    public static bool operator <=(EmptyCell a, dynamic b)
    {
        return 0 <= b;
    }

    public static bool operator <=(EmptyCell a, EmptyCell b)
    {
        return true;
    }

    public static bool operator >=(dynamic a, EmptyCell b)
    {
        return a >= 0;
    }

    public static bool operator >=(EmptyCell a, dynamic b)
    {
        return 0 >= b;
    }

    public static bool operator >=(EmptyCell a, EmptyCell b)
    {
        return true;
    }

    public static dynamic operator !(EmptyCell a)
    {
        return true;
    }

    public static implicit operator double(EmptyCell a) => 0;
    public static implicit operator string(EmptyCell a) => "";
    public static implicit operator bool(EmptyCell a) => false;

    public override bool Equals(object obj)
    {
        return obj == this;
    }

    public override int GetHashCode()
    {
        return 0;
    }
}