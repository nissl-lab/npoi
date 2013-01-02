using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Reflection;
using JavaToCSharp.Configuration;
using System.Configuration;
using JavaToCSharp.Rules;

namespace JavaToCSharp
{
    public class RuleEngine
    {
        List<Rule> rules = new List<Rule>();

        public void AddRule(Rule rule)
        {
            rules.Add(rule);
        }

        /// <summary>
        /// Loads the rules.
        /// </summary>
        public void LoadRules()
        {
            //load from app.config
            J2CSSection config =
                ConfigurationManager.GetSection(
                "J2CS") as J2CSSection;
            foreach (RuleSection rs in config.Rules)
            {
                EquivalentRule r = new EquivalentRule(rs.Name);
                r.Pattern = rs.Pattern;
                r.Replacement = rs.Replacement;
                //Console.WriteLine(r.ToString());
                AddRule(r);
            }
            //load from assembly
            Assembly asm = Assembly.GetExecutingAssembly();
            string nameSpace = "JavaToCSharp.Rules";
            foreach (Type type in asm.GetTypes())
            {
                if (type.Namespace == nameSpace && !type.IsAbstract && type.Name != "EquivalentRule")
                    AddRule(Activator.CreateInstance(type) as Rule);
            }
        }

        protected static void Log(string message)
        {
            Console.WriteLine(message);
        }
        protected static void LogTip(string message)
        {
            Console.WriteLine(message);
        }

        protected void LogExecute(string ruleName,int iRowNumber)
        {
            Log(string.Format("[Line {0}] {1}", iRowNumber, ruleName));
        }

        //protected void LogFind(string strOrigin, int iRowNumber)
        //{
        //    Log(string.Format("[{0}]", this.RuleName));
        //    Log(string.Format("(Line: {0}): {1}", iRowNumber, strOrigin.Trim().Replace("\r", "")));
        //}

        //protected void LogAttention(string strOrigin, int iRowNumber)
        //{
        //    LogTip(string.Format("[{0}]", this.RuleName));
        //    LogTip(string.Format("(Line: {0}): {1}", iRowNumber, strOrigin.Trim().Replace("\r", "")));
        //}
        public void Run(string path)
        {
            //read from file
            if (!File.Exists(path))
            {
                Console.WriteLine("File not found");
                return;
            }
            StreamReader sr = new StreamReader(path, true);
            string strOrigin = sr.ReadToEnd();
            sr.Close();

            //run rules
            StringBuilder sb = new StringBuilder();
            string[] arrInput = strOrigin.Split(new char[] { '\n' });
            for (int i = 0; i < arrInput.Length; i++)
            {
                string tmp = arrInput[i].Replace("\r", "");

                int ruleNum = i + 1;
                foreach (Rule rule in rules)
                {
                    if (rule.Execute(tmp, out tmp, ruleNum))
                    {
                        LogExecute(rule.RuleName, ruleNum);
                    }
                }
                sb.AppendLine(tmp);
            }

            //save result to file
            StreamWriter sw = new StreamWriter(path, false);
            sw.Write(sb.ToString());
            sw.Close();
        }

    }
}
