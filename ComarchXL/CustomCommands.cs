using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace ComarchXL
{
    // gdy chcemy utworzyć własne komendy o innych nazwach niż predefiniowane: New, Cut, Paste itd
    // musimu utworzyć własną klasę statyczną jak poniżej
    public static class CustomCommands
    {
        public static readonly RoutedUICommand Exit = new RoutedUICommand
            (
                "Exit",
                "Exit",
                typeof(CustomCommands),
                new InputGestureCollection()
                {
                    new KeyGesture(Key.F10, ModifierKeys.Alt)
                }
            );

        //Define more commands here, just like the one above
    }
}
