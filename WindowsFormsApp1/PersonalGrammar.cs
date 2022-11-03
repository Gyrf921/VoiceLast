using Microsoft.Speech.Recognition;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{
    class PersonalGrammar
    {
        static readonly CultureInfo _language = new CultureInfo("ru-RU");


        public static Grammar ChoiseGrammar()
        {
            GrammarBuilder gb_W = new GrammarBuilder();
            gb_W.Culture = _language;

            gb_W.Append("выбрать");
            gb_W.Append("папку");

            Grammar g_Weather = new Grammar(gb_W); //управляющий Grammar
            return g_Weather;
        }
        public static Grammar StandartPathGrammar()
        {
            GrammarBuilder gb_W = new GrammarBuilder();
            gb_W.Culture = _language;

            gb_W.Append("открыть");
            gb_W.Append("стартовую");
            gb_W.Append("папку");

            Grammar g_Weather = new Grammar(gb_W); //управляющий Grammar 
            return g_Weather;
        }
        public static Grammar GoToBackGrammar()
        {
            GrammarBuilder gb_W = new GrammarBuilder();
            gb_W.Culture = _language;

            gb_W.Append("выйти");
            gb_W.Append("из");
            gb_W.Append("папки");

            Grammar g_Weather = new Grammar(gb_W); //управляющий Grammar 
            return g_Weather;
        }


    }
}
