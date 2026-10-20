// 由 python 版 vba_builtins.py 自动导出，请勿手工编辑。
// 两边必须逐字一致：这是"词库里有什么"的唯一事实来源。

namespace VbeComplete
{
    internal static class BuiltinTables
    {
        // BUILTIN_FUNCTIONS: 171 条
        public static readonly string[] Functions =
        {
            "CBool", "CByte", "CCur", "CDate", "CDbl", "CDec",
            "CInt", "CLng", "CLngLng", "CLngPtr", "CSng", "CStr",
            "CVar", "CVErr", "IsArray", "IsDate", "IsEmpty", "IsError",
            "IsMissing", "IsNull", "IsNumeric", "IsObject", "TypeName", "VarType",
            "Asc", "AscB", "AscW", "Chr", "ChrB", "ChrW",
            "Format", "FormatCurrency", "FormatDateTime", "FormatNumber", "FormatPercent", "Hex",
            "Oct", "InStr", "InStrRev", "Join", "LCase", "Left",
            "LeftB", "Len", "LenB", "LTrim", "Mid", "MidB",
            "Replace", "Right", "RightB", "RTrim", "Space", "Split",
            "Str", "StrComp", "StrConv", "String", "StrReverse", "Trim",
            "UCase", "Val", "Filter", "Abs", "Atn", "Cos",
            "Exp", "Fix", "Int", "Log", "Rnd", "Round",
            "Sgn", "Sin", "Sqr", "Tan", "Date", "DateAdd",
            "DateDiff", "DatePart", "DateSerial", "DateValue", "Day", "Hour",
            "Minute", "Month", "MonthName", "Now", "Second", "Time",
            "Timer", "TimeSerial", "TimeValue", "Weekday", "WeekdayName", "Year",
            "Array", "LBound", "UBound", "MsgBox", "InputBox", "DoEvents",
            "Shell", "AppActivate", "SendKeys", "Command", "Environ", "CurDir",
            "CallByName", "Choose", "CreateObject", "GetObject", "IIf", "Partition",
            "QBColor", "RGB", "Switch", "Error", "Dir", "EOF",
            "FileAttr", "FileDateTime", "FileLen", "FreeFile", "GetAttr", "SetAttr",
            "FileCopy", "Kill", "MkDir", "RmDir", "ChDir", "ChDrive",
            "Open", "Close", "Get", "Put", "Input", "Seek",
            "Lock", "Unlock", "Loc", "LOF", "Print", "Write",
            "Reset", "LSet", "RSet", "LoadPicture", "SavePicture", "SaveSetting",
            "GetSetting", "GetAllSettings", "DeleteSetting", "DDB", "FV", "IPmt",
            "IRR", "MIRR", "NPer", "NPV", "Pmt", "PPmt",
            "PV", "Rate", "SLN", "SYD", "Load", "Unload",
            "Beep", "Erase", "Randomize",
        };

        // BUILTIN_TYPE_NAMES: 13 条
        public static readonly string[] TypeNames =
        {
            "Boolean", "Byte", "Currency", "Date", "Double", "Integer",
            "Long", "LongLong", "LongPtr", "Object", "Single", "String",
            "Variant",
        };

        // BUILTIN_KEYWORDS: 79 条
        public static readonly string[] Keywords =
        {
            "Sub", "Function", "Property", "Declare", "PtrSafe", "Alias",
            "Lib", "Public", "Private", "Friend", "Global", "Static",
            "Const", "Dim", "ReDim", "Preserve", "Implements", "Event",
            "WithEvents", "RaiseEvent", "Optional", "ParamArray", "ByVal", "ByRef",
            "Attribute", "Option", "Explicit", "Enum", "Type", "As",
            "New", "Me", "Is", "Like", "Set", "Let",
            "Call", "If", "Then", "Else", "ElseIf", "End",
            "Select", "Case", "For", "Each", "In", "To",
            "Step", "Next", "Do", "Loop", "While", "Wend",
            "Until", "Exit", "GoTo", "GoSub", "Return", "On",
            "Resume", "Stop", "With", "TypeOf", "Rem", "Debug",
            "True", "False", "Nothing", "Empty", "Null", "And",
            "Or", "Not", "Xor", "Eqv", "Imp", "Mod",
            "AddressOf",
        };

        // 补全后【要跟一个空格】的关键字（排除法算出来的，与 python 版同一份判据）
        public static readonly string[] KeywordsWantSpace =
        {
            "addressof", "alias", "and", "as", "attribute", "byref",
            "byval", "call", "case", "const", "declare", "dim",
            "each", "elseif", "end", "enum", "eqv", "event",
            "exit", "for", "friend", "function", "global", "gosub",
            "goto", "if", "imp", "implements", "in", "is",
            "let", "lib", "like", "mod", "new", "not",
            "on", "option", "optional", "or", "paramarray", "preserve",
            "private", "property", "ptrsafe", "public", "raiseevent", "redim",
            "rem", "select", "set", "static", "step", "sub",
            "to", "type", "typeof", "until", "while", "with",
            "withevents", "xor",
        };
    }
}
