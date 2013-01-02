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

using System.IO;
namespace NPOI.Util
{


/**
 * Build a 'lite' version of the ooxml-schemas.jar
 *
 * @author Yegor Kozlov
 */
public class OOXMLLite {

    private static Field _classes;
    static OOXMLLite(){
            //_classes = ClassLoader.class.GetDeclaredField("classes");
            _classes.SetAccessible(true);

    }

    /**
     * Destination directory to copy filtered classes
     */
    private File _destDest;

    /**
     * Directory with the compiled ooxml tests
     */
    private File _testDir;

    /**
     * Reference to the ooxml-schemas.jar
     */
    private Stream _ooxmlJar;


    OOXMLLite(String dest, String test, String ooxmlJar) {
        _destDest = new FileStream(dest);
        _testDir = new FileStream(test);
        _ooxmlJar = new FileStream(ooxmlJar);
    }

    public static void main(String[] args)  {

        String dest = null, test = null, ooxml = null;

        for (int i = 0; i < args.Length; i++) {
            if (args[i].Equals("-dest")) dest = args[++i];
            else if (args[i].Equals("-test")) test = args[++i];
            else if (args[i].Equals("-ooxml")) ooxml = args[++i];
        }
        OOXMLLite builder = new OOXMLLite(dest, test, ooxml);
        builder.build();
    }

    void build() {

        List<String> lst = new ArrayList<String>();
        //collect unit tests
        Console.WriteLine("Collecting unit tests from " + _testDir);
        collectTests(_testDir, _testDir, lst, ".+?\\.Test.+?\\.class$");

        TestSuite suite = new TestSuite();
        for (String arg : lst) {
            //ignore inner classes defined in tests
            if (arg.IndexOf('$') != -1) continue;

            String cls = arg.Replace(".class", "");
            try {
                Class test = Class.forName(cls);
                suite.AddTestSuite(test);
            } catch (ClassNotFoundException e) {
                throw new RuntimeException(e);
            }
        }

        //run tests
        TestRunner.Run(suite);

        //see what classes from the ooxml-schemas.jar are loaded
        Console.WriteLine("Copying classes to " + _destDest);
        Map<String, Class<?>> classes = GetLoadedClasses(_ooxmlJar.getName());
        for (Class<?> cls : classes.values()) {
            String className = cls.GetName();
            String classRef = className.Replace('.', '/') + ".class";
            File destFile = new File(_destDest, classRef);
            copyFile(cls.GetResourceAsStream('/' + classRef), destFile);

            if(cls.isInterface()){
                /**
                 * Copy classes and interfaces declared as members of this class
                 */
                for(Class fc : cls.GetDeclaredClasses()){
                    className = fc.GetName();
                    classRef = className.Replace('.', '/') + ".class";
                    destFile = new File(_destDest, classRef);
                    copyFile(fc.GetResourceAsStream('/' + classRef), destFile);
                }
            }
        }

        //finally copy the compiled .xsb files
        Console.WriteLine("Copying .xsb resources");
        JarFile jar = new  JarFile(_ooxmlJar);
        for(Enumeration<JarEntry> e = jar.entries(); e.hasMoreElements(); ){
            JarEntry je = e.nextElement();
            if(je.GetName().matches("schemaorg_apache_xmlbeans/system/\\w+/\\w+\\.xsb")) {
                 File destFile = new File(_destDest, je.GetName());
                 copyFile(jar.GetInputStream(je), destFile);
            }
        }
        jar.close();
    }

    /**
     * Recursively collect classes from the supplied directory
     *
     * @param arg   the directory to search in
     * @param out   output
     * @param ptrn  the pattern (regexp) to filter found files
     */
    private static void collectTests(File root, File arg, List<String> out, String ptrn) {
        if (arg.isDirectory()) {
            for (File f : arg.listFiles()) {
                collectTests(root, f, out, ptrn);
            }
        } else {
            String path = arg.GetAbsolutePath();
            String prefix = root.GetAbsolutePath();
            String cls = path.Substring(prefix.Length + 1).Replace(File.separator, ".");
            if(cls.matches(ptrn)) out.Add(cls);
        }
    }

    /**
     *
     * @param ptrn the pattern to filter output 
     * @return the classes loaded by the system class loader keyed by class name
     */
    @SuppressWarnings("unChecked")
    private static Map<String, Class<?>> GetLoadedClasses(String ptrn) {
        ClassLoader appLoader = ClassLoader.GetSystemClassLoader();
        try {
            Vector<Class<?>> classes = (Vector<Class<?>>) _classes.Get(appLoader);
            Map<String, Class<?>> map = new HashMap<String, Class<?>>();
            for (Class<?> cls : classes) {
                String jar = cls.GetProtectionDomain().getCodeSource().getLocation().toString();
                if(jar.IndexOf(ptrn) != -1) map.Put(cls.GetName(), cls);
            }
            return map;
        } catch (IllegalAccessException e) {
            throw new RuntimeException(e);
        }
    }

    private static void copyFile(InputStream srcStream, File destFile)  {
        File destDirectory = destFile.GetParentFile();
        destDirectory.mkdirs();
        OutputStream destStream = new FileOutputStream(destFile);
        try {
            IOUtils.copy(srcStream, destStream);
        } finally {
            destStream.close();
        }
    }

}
