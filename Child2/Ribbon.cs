using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using ExcelDna.Integration.Extensibility;
using System;
using System.Diagnostics;
using System.IO;

namespace Main
{
    // This class will load as a separate add-in
    public class Ribbon : ExcelRibbon
    {
        public override string GetCustomUI(string RibbonID)
        {
            // Return the Office customUI xml to add s ingle tab, group and button
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
                        <ribbon>
                            <tabs>
                                <tab id='Tab1' label='Child2'>
                                    <group id='Group1' label='Group1'>
                                        <button id='Button1' label='Button1' onAction='Button1_Click' />
                                    </group>
                                </tab>
                            </tabs>
                        </ribbon>
                    </customUI>";

        }

        public override void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
        }
        
        public override void OnStartupComplete(ref Array custom)
        {
            base.OnStartupComplete(ref custom);
        }

        // Implement the callback for the button
        public void Button1_Click(IRibbonControl _)
        {
            Debug.Print("Main Ribbon Button1_Click");
        }
    }
}