using cdn_api;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace ComarchXL
{
    public static class ComarchTools
    {
        public static int  zaloguj(ref int SessionId)
        {
            var Login = new XLLoginInfo_20201
            {
                Wersja = 20201,
                ProgramID = "test",
                OpeIdent = "admin",
                OpeHaslo = "1111",
                Baza = "KABAT_TEST"

            };
            
            cdn_api.cdn_api.XLLogin(Login, ref SessionId);
            MessageBox.Show(SessionId.ToString());
            
            return SessionId;
        }

        public static int wyloguj(int SessionId)
        {
            int retValue = cdn_api.cdn_api.XLLogout(SessionId);
            MessageBox.Show(retValue.ToString());
            return retValue;
        }


    }
}
