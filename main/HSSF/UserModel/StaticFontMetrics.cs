/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) Under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You Under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed Under the License is distributed on an "AS Is" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations Under the License.
==================================================================== */

namespace NPOI.HSSF.UserModel
{
    using System;
    using System.Collections;
    using System.IO;
    using System.Drawing;
    using System.Configuration;
    using NPOI.Util.Collections;

    /**
     * Allows the user to lookup the font metrics for a particular font without
     * actually having the font on the system.  The font details are Loaded
     * as a resource from the POI jar file (or classpath) and should be contained
     * in path "/font_metrics.properties".  The font widths are for a 10 point
     * version of the font.  Use a multiplier for other sizes.
     *
     * @author Glen Stampoultzis (glens at apache.org)
     */
    class StaticFontMetrics
    {
#if NETSTANDARD2_1 || NETSTANDARD2_0
        private const String FONT_METRICS_PROPERTIES_FILE_NAME = "NPOI.Resources.font_metrics.properties";
#else
        private const String FONT_METRICS_PROPERTIES_FILE_NAME = "font_metrics.properties";
#endif
        
        /** The font metrics property file we're using */
        private static Properties fontMetricsProps;
        /** Our cache of font details we've alReady looked up */
        private static Hashtable fontDetailsMap = new Hashtable();

        /**
         * Retrieves the fake font details for a given font.
         * @param font  the font to lookup.
         * @return  the fake font.
         */
        public static FontDetails GetFontDetails(Font font)
        {
            // If we haven't alReady identified out font metrics file,
            //  figure out which one to use and Load it
            if (fontMetricsProps == null)
            {
                Stream metricsIn = null;
                try
                {
                    fontMetricsProps = new Properties();

                    // Check to see if the font metric file was specified
                    //  as a system property
                    String propFileName = null;
                    try
                    {
                        propFileName = ConfigurationManager.AppSettings["font.metrics.filename"];
                    }
                    catch(Exception) { }

                    if (propFileName != null)
                    {
                        
                        if (!File.Exists(propFileName))
                            throw new FileNotFoundException("font_metrics.properties not found at path " + Path.GetFullPath(propFileName));
                        metricsIn = typeof (StaticFontMetrics).Assembly.GetManifestResourceStream (FONT_METRICS_PROPERTIES_FILE_NAME);
                    }
                    else
                    {
                        // Use the built-in font metrics file off the classpath
                        metricsIn = typeof (StaticFontMetrics).Assembly.GetManifestResourceStream (FONT_METRICS_PROPERTIES_FILE_NAME);
                        if (metricsIn == null)
                            throw new FileNotFoundException("font_metrics.properties not found in classpath");
                    }
                    fontMetricsProps.Load(metricsIn);
                }
                catch (IOException e)
                {
                    throw new Exception("Could not Load font metrics: " + e.Message);
                }
                finally
                {
                    if (metricsIn != null)
                    {
                        try
                        {
                            metricsIn.Close();
                        }
                        catch (IOException)
                        {
                        
                        }
                    }
                }
            }

            // Grab the base name of the font they've asked about
            String fontName = font.FontFamily.Name;

            // Some fonts support plain/bold/italic/bolditalic variants
            // Others have different font instances for bold etc
            // (eg font.dialog.plain.* vs font.Californian FB Bold.*)
            String fontStyle = "";
            //if(font.IsPlain())  fontStyle += "plain";
            if (font.Bold) fontStyle += "bold";
            if (font.Italic) fontStyle += "italic";

            // Do we have a definition for this font with just the name?
            // If not, Check with the font style Added
            if (fontMetricsProps[FontDetails.BuildFontHeightProperty(fontName)] == null &&
            fontMetricsProps[FontDetails.BuildFontHeightProperty(fontName + "." + fontStyle)] != null)
            {
                // Need to Add on the style to the font name
                fontName += "." + fontStyle;
            }

            // Get the details on this font
            if (fontDetailsMap[fontName] == null)
            {
                FontDetails fontDetails = FontDetails.Create(fontName, fontMetricsProps);
                fontDetailsMap[fontName]= fontDetails;
                return fontDetails;
            }
            else
            {
                return (FontDetails)fontDetailsMap[fontName];
            }

        }
    }
}