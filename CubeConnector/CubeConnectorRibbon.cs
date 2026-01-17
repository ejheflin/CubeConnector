/*
 * CubeConnector - Excel-DNA add-in for querying Power BI datasets
 * Copyright (C) 2026
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program. If not, see <https://www.gnu.org/licenses/>.
 *
 * For enterprise licensing options, please contact the project maintainers.
 */

using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using System;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;

namespace CubeConnector
{
    [ComVisible(true)]
    public class CubeConnectorRibbon : ExcelRibbon
    {
        public override string GetCustomUI(string RibbonID)
        {
            return @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' loadImage='LoadImage'>
    <ribbon>
        <tabs>
            <tab idMso='TabData'>
                <group id='CubeConnectorGroup' label='CubeConnector' insertAfterMso='GroupRefreshAll'>
                    <splitButton id='RefreshCCSplitButton' size='large'>
                        <button id='RefreshCCButton' 
                                label='Refresh'
                                onAction='OnRefreshClicked'
                                image='CubeIcon' />
                        <menu id='RefreshCCMenu'>
                            <button id='RefreshCacheBtn' 
                                    label='Refresh Cache' 
                                    onAction='OnRefreshCacheClicked'
                                    imageMso='RefreshAll' />
                            <button id='DrillToDetailsBtn' 
                                    label='Drill to Details' 
                                    onAction='OnDrillToDetailsClicked'
                                    imageMso='ControlWizards' />
                            <button id='DrillToPivotBtn' 
                                    label='Drill to Pivot' 
                                    onAction='OnDrillToPivotClicked'
                                    imageMso='PivotTableInsert' />
                        </menu>
                    </splitButton>
                </group>
            </tab>
        </tabs>
    </ribbon>
</customUI>";
        }

        public override object LoadImage(string imageId)
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                string resourceName = "CubeConnector.cubeconnectorlogo.png";

                var stream = assembly.GetManifestResourceStream(resourceName);
                if (stream != null)
                {
                    var bitmap = new Bitmap(stream);
                    return AxHostConverter.ImageToIPictureDisp(bitmap);
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error loading image: {ex.Message}");
            }
            return null;
        }

        public void OnRefreshClicked(IRibbonControl control)
        {
            try
            {
                // Ensure prerequisites exist
                EnsureConnectionExists();
                EnsureCacheExists();

                var app = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
                var workbook = app.ActiveWorkbook;
                var manager = new RefreshManager(app, workbook);
                manager.RefreshAll();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    $"Error refreshing cache:\n\n{ex.Message}",
                    "CubeConnector Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        public void OnRefreshCacheClicked(IRibbonControl control)
        {
            // Ensure prerequisites exist
            EnsureConnectionExists();
            EnsureCacheExists();
            // Same as main button
            OnRefreshClicked(control);
        }

        public void OnDrillToDetailsClicked(IRibbonControl control)
        {
            DynamicFunctionRegistration.DrillToDetailsHandler();
        }

        public void OnDrillToPivotClicked(IRibbonControl control)
        {
            DynamicFunctionRegistration.DrillToPivotHandler();
        }
        private static void EnsureConnectionExists()
        {
            DynamicFunctionRegistration.EnsureConnectionExists();
        }

        private static void EnsureCacheExists()
        {
            DynamicFunctionRegistration.EnsureCacheExists();
        }
    }

    // Helper class to access protected GetIPictureDispFromPicture
    internal class AxHostConverter : System.Windows.Forms.AxHost
    {
        private AxHostConverter() : base(string.Empty) { }

        public static object ImageToIPictureDisp(System.Drawing.Image image)
        {
            return GetIPictureDispFromPicture(image);
        }
    }
}