using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;


namespace PriceList
{
    class clsMenuEvent
    {
        SAPbouiCOM.Form objform;

        public void MenuEvent_For_StandardMenu(ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                switch (clsModule.objaddon.objapplication.Forms.ActiveForm.TypeEx)
                {
                    case "PriceList":
                        // {
                        //   if (pVal.BeforeAction == true)
                        //     return;
                        //objform = clsModule.objaddon.objapplication.Forms.ActiveForm;
                        //Default_Sample_MenuEvent(pVal, BubbleEvent);


                        //}
                        break;
                }
            }
            catch (Exception ex)
            {

            }
        }
    }
}
