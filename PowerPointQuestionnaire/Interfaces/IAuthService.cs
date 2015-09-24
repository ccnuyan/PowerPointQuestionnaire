using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace PowerPointQuestionnaire.Interfaces
{
    public interface IAuthService
    {
        
        Task<bool> Authenticate(string username, string password);


        HttpWebRequest AddToken(HttpWebRequest httpWebRequest);
    }
}
