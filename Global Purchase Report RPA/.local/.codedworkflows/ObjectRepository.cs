using UiPath.CodedWorkflows.DescriptorIntegration;

namespace GlobalPurchaseReport.ObjectRepository
{
    public static class Descriptors
    {
        public static class __Chrome__Login
        {
            static string _reference = "gZXR734txESI5U3Knb0mjA/J8B6dV0LPkG0vVwM2ORztA";
            public static _Implementation.___Chrome__Login.__Chrome__Login Chrome__Login { get; private set; } = new _Implementation.___Chrome__Login.__Chrome__Login();
        }

        public static class Chrome__Login_1_
        {
            static string _reference = "gZXR734txESI5U3Knb0mjA/4iize1385UShdUjlWSoeKQ";
            public static _Implementation._Chrome__Login_1_.__Chrome__Login Chrome__Login { get; private set; } = new _Implementation._Chrome__Login_1_.__Chrome__Login();
        }

        public static class Chrome__Login_2_
        {
            static string _reference = "gZXR734txESI5U3Knb0mjA/p6ee9bxJt0aOYi1lulLV-w";
            public static _Implementation._Chrome__Login_2_.__Chrome__Login Chrome__Login { get; private set; } = new _Implementation._Chrome__Login_2_.__Chrome__Login();
        }

        public static class Chrome__Login_3_
        {
            static string _reference = "gZXR734txESI5U3Knb0mjA/p580Ye911kq9Ilkrl_vzGA";
            public static _Implementation._Chrome__Login_3_.__Chrome__Login Chrome__Login { get; private set; } = new _Implementation._Chrome__Login_3_.__Chrome__Login();
        }

        public static class Chrome__Login_4_
        {
            static string _reference = "gZXR734txESI5U3Knb0mjA/4135QdM9Akuj2HaoJsmtMw";
            public static _Implementation._Chrome__Login_4_.__Chrome__Login Chrome__Login { get; private set; } = new _Implementation._Chrome__Login_4_.__Chrome__Login();
        }

        public static class __Chrome__Reports___MEYER_EPICOR
        {
            static string _reference = "gZXR734txESI5U3Knb0mjA/idpf_yuRx0aBkpZP66hZbg";
            public static _Implementation.___Chrome__Reports___MEYER_EPICOR.__Chrome__Reports___MEYER_EPICOR Chrome__Reports___MEYER_EPICOR { get; private set; } = new _Implementation.___Chrome__Reports___MEYER_EPICOR.__Chrome__Reports___MEYER_EPICOR();
        }
    }
}

namespace GlobalPurchaseReport._Implementation
{
    internal class ScreenDescriptorDefinition : IScreenDescriptorDefinition
    {
        public IScreenDescriptor Screen { get; set; }
        public string Reference { get; set; }
        public string DisplayName { get; set; }
    }

    internal class ElementDescriptorDefinition : IElementDescriptorDefinition
    {
        public IScreenDescriptor Screen { get; set; }
        public string Reference { get; set; }
        public string DisplayName { get; set; }
        public IElementDescriptor ParentElement { get; set; }
        public IElementDescriptor Element { get; set; }
    }

    namespace ___Chrome__Login._Chrome__Login
    {
        public class __INPUT_EPICOR_SUB_Member_ : IElementDescriptor
        {
            private readonly IScreenDescriptor _screenDescriptor;
            private readonly IElementDescriptor _parentElementDescriptor;
            private readonly IElementDescriptorDefinition _elementDescriptor;

            public IElementDescriptorDefinition GetDefinition()
            {
                return _elementDescriptor;
            }

            public __INPUT_EPICOR_SUB_Member_(IScreenDescriptor screenDescriptor, IElementDescriptor parentElementDescriptor)
            {
                _screenDescriptor = screenDescriptor;
                _parentElementDescriptor = parentElementDescriptor;
                _elementDescriptor = new ElementDescriptorDefinition
                {
                    Reference = "gZXR734txESI5U3Knb0mjA/APWUAwLTekOLzAQuPrvIxw",
                    DisplayName = "INPUT EPICOR_SUB_Member…",
                    Element = this,
                    ParentElement = _parentElementDescriptor,
                    Screen = screenDescriptor
                };
            }
        }
    }

    namespace ___Chrome__Login._Chrome__Login
    {
        public class __INPUT_EPICOR_SUB_Member__1_ : IElementDescriptor
        {
            private readonly IScreenDescriptor _screenDescriptor;
            private readonly IElementDescriptor _parentElementDescriptor;
            private readonly IElementDescriptorDefinition _elementDescriptor;

            public IElementDescriptorDefinition GetDefinition()
            {
                return _elementDescriptor;
            }

            public __INPUT_EPICOR_SUB_Member__1_(IScreenDescriptor screenDescriptor, IElementDescriptor parentElementDescriptor)
            {
                _screenDescriptor = screenDescriptor;
                _parentElementDescriptor = parentElementDescriptor;
                _elementDescriptor = new ElementDescriptorDefinition
                {
                    Reference = "gZXR734txESI5U3Knb0mjA/t1DLyE9hmkOMu3amn4Glsw",
                    DisplayName = "INPUT EPICOR_SUB_Member…(1)",
                    Element = this,
                    ParentElement = _parentElementDescriptor,
                    Screen = screenDescriptor
                };
            }
        }
    }

    namespace ___Chrome__Login._Chrome__Login
    {
        public class __Sign_In : IElementDescriptor
        {
            private readonly IScreenDescriptor _screenDescriptor;
            private readonly IElementDescriptor _parentElementDescriptor;
            private readonly IElementDescriptorDefinition _elementDescriptor;

            public IElementDescriptorDefinition GetDefinition()
            {
                return _elementDescriptor;
            }

            public __Sign_In(IScreenDescriptor screenDescriptor, IElementDescriptor parentElementDescriptor)
            {
                _screenDescriptor = screenDescriptor;
                _parentElementDescriptor = parentElementDescriptor;
                _elementDescriptor = new ElementDescriptorDefinition
                {
                    Reference = "gZXR734txESI5U3Knb0mjA/TDefh7hGxUyKgsCoSi50IA",
                    DisplayName = "Sign In",
                    Element = this,
                    ParentElement = _parentElementDescriptor,
                    Screen = screenDescriptor
                };
            }
        }
    }

    namespace ___Chrome__Login
    {
        public class __Chrome__Login : IScreenDescriptor
        {
            public IScreenDescriptorDefinition GetDefinition()
            {
                return _screenDescriptor;
            }

            private readonly ScreenDescriptorDefinition _screenDescriptor;

            public __Chrome__Login()
            {
                _screenDescriptor = new ScreenDescriptorDefinition
                {
                    Reference = "gZXR734txESI5U3Knb0mjA/uMAGWxX2cEmn580B8F_QQQ",
                    DisplayName = "Chrome: Login",
                    Screen = this
                };
                INPUT_EPICOR_SUB_Member_ = new _Implementation.___Chrome__Login._Chrome__Login.__INPUT_EPICOR_SUB_Member_(this, null);
                INPUT_EPICOR_SUB_Member__1_ = new _Implementation.___Chrome__Login._Chrome__Login.__INPUT_EPICOR_SUB_Member__1_(this, null);
                Sign_In = new _Implementation.___Chrome__Login._Chrome__Login.__Sign_In(this, null);
            }

            public _Implementation.___Chrome__Login._Chrome__Login.__INPUT_EPICOR_SUB_Member_ INPUT_EPICOR_SUB_Member_ { get; private set; }
            public _Implementation.___Chrome__Login._Chrome__Login.__INPUT_EPICOR_SUB_Member__1_ INPUT_EPICOR_SUB_Member__1_ { get; private set; }
            public _Implementation.___Chrome__Login._Chrome__Login.__Sign_In Sign_In { get; private set; }
        }
    }

    namespace _Chrome__Login_1_._Chrome__Login
    {
        public class __Report : IElementDescriptor
        {
            private readonly IScreenDescriptor _screenDescriptor;
            private readonly IElementDescriptor _parentElementDescriptor;
            private readonly IElementDescriptorDefinition _elementDescriptor;

            public IElementDescriptorDefinition GetDefinition()
            {
                return _elementDescriptor;
            }

            public __Report(IScreenDescriptor screenDescriptor, IElementDescriptor parentElementDescriptor)
            {
                _screenDescriptor = screenDescriptor;
                _parentElementDescriptor = parentElementDescriptor;
                _elementDescriptor = new ElementDescriptorDefinition
                {
                    Reference = "gZXR734txESI5U3Knb0mjA/TMyuxe_CCU-8Tl67140gdQ",
                    DisplayName = "Report",
                    Element = this,
                    ParentElement = _parentElementDescriptor,
                    Screen = screenDescriptor
                };
            }
        }
    }

    namespace _Chrome__Login_1_
    {
        public class __Chrome__Login : IScreenDescriptor
        {
            public IScreenDescriptorDefinition GetDefinition()
            {
                return _screenDescriptor;
            }

            private readonly ScreenDescriptorDefinition _screenDescriptor;

            public __Chrome__Login()
            {
                _screenDescriptor = new ScreenDescriptorDefinition
                {
                    Reference = "gZXR734txESI5U3Knb0mjA/OO3nBLD5-ka59luryS9FmA",
                    DisplayName = "Chrome: Login",
                    Screen = this
                };
                Report = new _Implementation._Chrome__Login_1_._Chrome__Login.__Report(this, null);
            }

            public _Implementation._Chrome__Login_1_._Chrome__Login.__Report Report { get; private set; }
        }
    }

    namespace _Chrome__Login_2_._Chrome__Login
    {
        public class __Report : IElementDescriptor
        {
            private readonly IScreenDescriptor _screenDescriptor;
            private readonly IElementDescriptor _parentElementDescriptor;
            private readonly IElementDescriptorDefinition _elementDescriptor;

            public IElementDescriptorDefinition GetDefinition()
            {
                return _elementDescriptor;
            }

            public __Report(IScreenDescriptor screenDescriptor, IElementDescriptor parentElementDescriptor)
            {
                _screenDescriptor = screenDescriptor;
                _parentElementDescriptor = parentElementDescriptor;
                _elementDescriptor = new ElementDescriptorDefinition
                {
                    Reference = "gZXR734txESI5U3Knb0mjA/ARYKVo0iEUSr6F6uW27aCw",
                    DisplayName = "Report",
                    Element = this,
                    ParentElement = _parentElementDescriptor,
                    Screen = screenDescriptor
                };
            }
        }
    }

    namespace _Chrome__Login_2_
    {
        public class __Chrome__Login : IScreenDescriptor
        {
            public IScreenDescriptorDefinition GetDefinition()
            {
                return _screenDescriptor;
            }

            private readonly ScreenDescriptorDefinition _screenDescriptor;

            public __Chrome__Login()
            {
                _screenDescriptor = new ScreenDescriptorDefinition
                {
                    Reference = "gZXR734txESI5U3Knb0mjA/JvDzOtvAEESy0qcYPZpdww",
                    DisplayName = "Chrome: Login",
                    Screen = this
                };
                Report = new _Implementation._Chrome__Login_2_._Chrome__Login.__Report(this, null);
            }

            public _Implementation._Chrome__Login_2_._Chrome__Login.__Report Report { get; private set; }
        }
    }

    namespace _Chrome__Login_3_._Chrome__Login
    {
        public class __INPUT_dp1759463294720 : IElementDescriptor
        {
            private readonly IScreenDescriptor _screenDescriptor;
            private readonly IElementDescriptor _parentElementDescriptor;
            private readonly IElementDescriptorDefinition _elementDescriptor;

            public IElementDescriptorDefinition GetDefinition()
            {
                return _elementDescriptor;
            }

            public __INPUT_dp1759463294720(IScreenDescriptor screenDescriptor, IElementDescriptor parentElementDescriptor)
            {
                _screenDescriptor = screenDescriptor;
                _parentElementDescriptor = parentElementDescriptor;
                _elementDescriptor = new ElementDescriptorDefinition
                {
                    Reference = "gZXR734txESI5U3Knb0mjA/hRr0SgEpP0S0j1Y5HBarEg",
                    DisplayName = "INPUT dp1759463294720",
                    Element = this,
                    ParentElement = _parentElementDescriptor,
                    Screen = screenDescriptor
                };
            }
        }
    }

    namespace _Chrome__Login_3_._Chrome__Login
    {
        public class __INPUT_EPICOR_SUB_Report_ : IElementDescriptor
        {
            private readonly IScreenDescriptor _screenDescriptor;
            private readonly IElementDescriptor _parentElementDescriptor;
            private readonly IElementDescriptorDefinition _elementDescriptor;

            public IElementDescriptorDefinition GetDefinition()
            {
                return _elementDescriptor;
            }

            public __INPUT_EPICOR_SUB_Report_(IScreenDescriptor screenDescriptor, IElementDescriptor parentElementDescriptor)
            {
                _screenDescriptor = screenDescriptor;
                _parentElementDescriptor = parentElementDescriptor;
                _elementDescriptor = new ElementDescriptorDefinition
                {
                    Reference = "gZXR734txESI5U3Knb0mjA/tL4pM3IoVkqe_y1WNxiceQ",
                    DisplayName = "INPUT EPICOR_SUB_Report…",
                    Element = this,
                    ParentElement = _parentElementDescriptor,
                    Screen = screenDescriptor
                };
            }
        }
    }

    namespace _Chrome__Login_3_._Chrome__Login
    {
        public class __UL : IElementDescriptor
        {
            private readonly IScreenDescriptor _screenDescriptor;
            private readonly IElementDescriptor _parentElementDescriptor;
            private readonly IElementDescriptorDefinition _elementDescriptor;

            public IElementDescriptorDefinition GetDefinition()
            {
                return _elementDescriptor;
            }

            public __UL(IScreenDescriptor screenDescriptor, IElementDescriptor parentElementDescriptor)
            {
                _screenDescriptor = screenDescriptor;
                _parentElementDescriptor = parentElementDescriptor;
                _elementDescriptor = new ElementDescriptorDefinition
                {
                    Reference = "gZXR734txESI5U3Knb0mjA/EvbWcjliWUWKxdplCtT8yA",
                    DisplayName = "UL",
                    Element = this,
                    ParentElement = _parentElementDescriptor,
                    Screen = screenDescriptor
                };
            }
        }
    }

    namespace _Chrome__Login_3_
    {
        public class __Chrome__Login : IScreenDescriptor
        {
            public IScreenDescriptorDefinition GetDefinition()
            {
                return _screenDescriptor;
            }

            private readonly ScreenDescriptorDefinition _screenDescriptor;

            public __Chrome__Login()
            {
                _screenDescriptor = new ScreenDescriptorDefinition
                {
                    Reference = "gZXR734txESI5U3Knb0mjA/YxDMQ823oEa-L_pHjEXrQQ",
                    DisplayName = "Chrome: Login",
                    Screen = this
                };
                INPUT_dp1759463294720 = new _Implementation._Chrome__Login_3_._Chrome__Login.__INPUT_dp1759463294720(this, null);
                INPUT_EPICOR_SUB_Report_ = new _Implementation._Chrome__Login_3_._Chrome__Login.__INPUT_EPICOR_SUB_Report_(this, null);
                UL = new _Implementation._Chrome__Login_3_._Chrome__Login.__UL(this, null);
            }

            public _Implementation._Chrome__Login_3_._Chrome__Login.__INPUT_dp1759463294720 INPUT_dp1759463294720 { get; private set; }
            public _Implementation._Chrome__Login_3_._Chrome__Login.__INPUT_EPICOR_SUB_Report_ INPUT_EPICOR_SUB_Report_ { get; private set; }
            public _Implementation._Chrome__Login_3_._Chrome__Login.__UL UL { get; private set; }
        }
    }

    namespace _Chrome__Login_4_._Chrome__Login
    {
        public class __Export : IElementDescriptor
        {
            private readonly IScreenDescriptor _screenDescriptor;
            private readonly IElementDescriptor _parentElementDescriptor;
            private readonly IElementDescriptorDefinition _elementDescriptor;

            public IElementDescriptorDefinition GetDefinition()
            {
                return _elementDescriptor;
            }

            public __Export(IScreenDescriptor screenDescriptor, IElementDescriptor parentElementDescriptor)
            {
                _screenDescriptor = screenDescriptor;
                _parentElementDescriptor = parentElementDescriptor;
                _elementDescriptor = new ElementDescriptorDefinition
                {
                    Reference = "gZXR734txESI5U3Knb0mjA/GjAMIOZfKEGy2CcPYuWSxQ",
                    DisplayName = "Export",
                    Element = this,
                    ParentElement = _parentElementDescriptor,
                    Screen = screenDescriptor
                };
            }
        }
    }

    namespace _Chrome__Login_4_
    {
        public class __Chrome__Login : IScreenDescriptor
        {
            public IScreenDescriptorDefinition GetDefinition()
            {
                return _screenDescriptor;
            }

            private readonly ScreenDescriptorDefinition _screenDescriptor;

            public __Chrome__Login()
            {
                _screenDescriptor = new ScreenDescriptorDefinition
                {
                    Reference = "gZXR734txESI5U3Knb0mjA/62G-Ln_Is0yUokfdC9-zpQ",
                    DisplayName = "Chrome: Login",
                    Screen = this
                };
                Export = new _Implementation._Chrome__Login_4_._Chrome__Login.__Export(this, null);
            }

            public _Implementation._Chrome__Login_4_._Chrome__Login.__Export Export { get; private set; }
        }
    }

    namespace ___Chrome__Reports___MEYER_EPICOR._Chrome__Reports___MEYER_EPICOR
    {
        public class __ACT001__ : IElementDescriptor
        {
            private readonly IScreenDescriptor _screenDescriptor;
            private readonly IElementDescriptor _parentElementDescriptor;
            private readonly IElementDescriptorDefinition _elementDescriptor;

            public IElementDescriptorDefinition GetDefinition()
            {
                return _elementDescriptor;
            }

            public __ACT001__(IScreenDescriptor screenDescriptor, IElementDescriptor parentElementDescriptor)
            {
                _screenDescriptor = screenDescriptor;
                _parentElementDescriptor = parentElementDescriptor;
                _elementDescriptor = new ElementDescriptorDefinition
                {
                    Reference = "gZXR734txESI5U3Knb0mjA/nEv_J5BZH0Shq3MPpQ0ipg",
                    DisplayName = "ACT001 -",
                    Element = this,
                    ParentElement = _parentElementDescriptor,
                    Screen = screenDescriptor
                };
            }
        }
    }

    namespace ___Chrome__Reports___MEYER_EPICOR
    {
        public class __Chrome__Reports___MEYER_EPICOR : IScreenDescriptor
        {
            public IScreenDescriptorDefinition GetDefinition()
            {
                return _screenDescriptor;
            }

            private readonly ScreenDescriptorDefinition _screenDescriptor;

            public __Chrome__Reports___MEYER_EPICOR()
            {
                _screenDescriptor = new ScreenDescriptorDefinition
                {
                    Reference = "gZXR734txESI5U3Knb0mjA/xtcpKqrqkEaKLB_QKve57g",
                    DisplayName = "Chrome: Reports - MEYER EPICOR",
                    Screen = this
                };
                ACT001__ = new _Implementation.___Chrome__Reports___MEYER_EPICOR._Chrome__Reports___MEYER_EPICOR.__ACT001__(this, null);
            }

            public _Implementation.___Chrome__Reports___MEYER_EPICOR._Chrome__Reports___MEYER_EPICOR.__ACT001__ ACT001__ { get; private set; }
        }
    }
}