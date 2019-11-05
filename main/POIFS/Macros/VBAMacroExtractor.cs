/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

namespace NPOI.POIFS.Macros
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Text;


    /**
     * This tool extracts out the source of all VBA Modules of an office file,
     *  both OOXML (eg XLSM) and OLE2/POIFS (eg DOC), to STDOUT or a directory.
     * 
     * @since 3.15-beta2
     */
    public class VBAMacroExtractor
    {
        public static void main(String[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Use:");
                Console.WriteLine("   VBAMacroExtractor <office.doc> [output]");
                Console.WriteLine("");
                Console.WriteLine("If an output directory is given, macros are written there");
                Console.WriteLine("Otherwise they are output to the screen");
                //System.Exit(1);
            }

            FileInfo input = new FileInfo(args[0]);
            DirectoryInfo output = null;
            if (args.Length > 1)
            {
                output = new DirectoryInfo(args[1]);
            }

            VBAMacroExtractor extractor = new VBAMacroExtractor();
            extractor.Extract(input, output);
        }

        /**
          * Extracts the VBA modules from a macro-enabled office file and Writes them
          * to files in <tt>outputDir</tt>.
          *
          * Creates the <tt>outputDir</tt>, directory, including any necessary but
          * nonexistent parent directories, if <tt>outputDir</tt> does not exist.
          * If <tt>outputDir</tt> is null, Writes the contents to standard out instead.
          *
          * @param input  the macro-enabled office file.
          * @param outputDir  the directory to write the extracted VBA modules to.
          * @param extension  file extension of the extracted VBA modules
          * @since 3.15-beta2
          */
        public void Extract(FileInfo input, DirectoryInfo outputDir, String extension)
        {
            if (!input.Exists) throw new FileNotFoundException(input.ToString());
            //System.err.Print("Extracting VBA Macros from " + input + " to ");
            if (outputDir != null)
            {
                if (!outputDir.Exists)
                {
                    outputDir.Create();
                    //throw new IOException("Output directory " + outputDir + " could not be Created");
                }
                Console.WriteLine(outputDir);
            }
            else
            {
                Console.WriteLine("STDOUT");
            }

            VBAMacroReader Reader = new VBAMacroReader(input);
            Dictionary<String, String> macros = Reader.ReadMacros();
            Reader.Close();

            String divider = "---------------------------------------";
            foreach (KeyValuePair<String, String> entry in macros)
            {
                String moduleName = entry.Key;
                String moduleCode = entry.Value;
                if (outputDir == null)
                {
                    Console.WriteLine(divider);
                    Console.WriteLine(moduleName);
                    Console.WriteLine("");
                    Console.WriteLine(moduleCode);
                }
                else
                {
                    FileInfo out1 = new FileInfo(Path.Combine(outputDir.FullName, moduleName + extension));
                    FileStream fout = out1.Create();
                    StreamWriter fWriter = new StreamWriter(fout, Encoding.UTF8);
                    fWriter.Write(moduleCode);
                    fWriter.Close();
                    fout.Close();
                    Console.WriteLine("Extracted " + out1);
                }
            }
            if (outputDir == null)
            {
                //System.out.Println(divider);
            }
        }

        /**
          * Extracts the VBA modules from a macro-enabled office file and Writes them
          * to <tt>.vba</tt> files in <tt>outputDir</tt>. 
          * 
          * Creates the <tt>outputDir</tt>, directory, including any necessary but
          * nonexistent parent directories, if <tt>outputDir</tt> does not exist.
          * If <tt>outputDir</tt> is null, Writes the contents to standard out instead.
          * 
          * @param input  the macro-enabled office file.
          * @param outputDir  the directory to write the extracted VBA modules to.
          * @since 3.15-beta2
          */
        public void Extract(FileInfo input, DirectoryInfo outputDir)
        {
            Extract(input, outputDir, ".vba");
        }
    }

}