namespace SCADAutoRunner
{
    /// <summary>
    /// Win32 constants
    /// </summary>
    static class Win32
    {
        // The WM_COMMAND message is sent when the user selects a command item from 
        // a menu, when a control sends a notification message to its parent window, 
        // or when an accelerator keystroke is translated.
        public const int WM_CLOSE = 0x10;
        public const int WM_KEYDOWN = 0x100;
        public const int WM_KEYUP = 0x101;
        public const int WM_COMMAND = 0x111;
        public const int WM_LBUTTONDOWN = 0x201;
        public const int WM_LBUTTONUP = 0x202;
        public const int WM_LBUTTONDBLCLK = 0x203;
        public const int WM_RBUTTONDOWN = 0x204;
        public const int WM_RBUTTONUP = 0x205;
        public const int WM_RBUTTONDBLCLK = 0x206;
        public const int WM_SETTEXT = 0x000c;

        public const int VK_TAB = 0x09;
        public const int VK_RETURN = 0x0D;
        public const int VK_DOWN = 0x28;
        public const int VK_SPACE = 0x20;
    }
}
