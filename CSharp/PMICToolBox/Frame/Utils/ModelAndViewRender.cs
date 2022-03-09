//------------------------------------------------------------------------------
// Copyright (C) 2018 Teradyne, Inc. All rights reserved.
//
// This document contains proprietary and confidential information of Teradyne,
// Inc. and is tendered subject to the condition that the information (a) be
// retained in confidence and (b) not be used or incorporated in any product
// except with the express written consent of Teradyne, Inc.
//
// <File description paragraph>
//
// Revision History:
// (Place the most recent revision history at the top.)
// Date        Name           Bug#            Notes
//
// 2018 Feb 28 Oliver Ou                      Initial creation
//
//------------------------------------------------------------------------------ 

using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace FWFrame.Utils
{
    public class ModelAndViewRender
    {
        public static Dictionary<string, List<string>> tempFileBuff = new Dictionary<string, List<string>>();

        const string KEY_GROUP_START = "{GROUP_START}";
        const string KEY_GROUP_END = "{GROUP_END}";
        const string KEY_LIST = "{LIST_";

        public static List<string> GetTmpLines(string templateFilePath)
        {
            if(tempFileBuff.Keys.Contains(templateFilePath))
            {
                return tempFileBuff[templateFilePath];
            }
            else
            {
                List<string> tempLines = File.ReadAllLines(templateFilePath).ToList();
                tempFileBuff.Add(templateFilePath, tempLines);

                return tempLines;
            }
        }

        /// <summary>
        /// replace key string({GROUP_START}, {GROUP_END}) in template file with data in model to generate result file.
        /// </summary>
        /// <param name="model">The model.</param>
        /// <param name="templateFilePath">The template file path.</param>
        /// <param name="resultFilePath">The result file path.</param>
        public static void RenderToFile(Dictionary<string, object> model, string templateFilePath, string resultFilePath)
        {
            List<string> tmpRenderLines = RenderToList(model, templateFilePath);

            // remove null or blank lines at end of string list
            bool blankStart = false;
            for (int i = tmpRenderLines.Count - 1; i >= 0; i--)
            {
                if (tmpRenderLines[i].IsNullOrBlank())
                {
                    if (!blankStart)
                    {
                        blankStart = true;
                        tmpRenderLines[i] = string.Empty;
                    }
                    else
                    {
                        tmpRenderLines.RemoveAt(i);
                    }
                }
                else
                {
                    blankStart = false;
                }
            }

            File.WriteAllLines(resultFilePath, tmpRenderLines);
        }

        /// <summary>
        /// read from template file and replace key string with data in model
        /// </summary>
        /// <param name="model">The model.</param>
        /// <param name="templateFilePath">The template file path.</param>
        /// <returns></returns>
        /// <exception cref="FWFrameException"></exception>
        public static List<string> RenderToList(Dictionary<string, object> model, string templateFilePath)
        {
            List<string> tempLines = GetTmpLines(templateFilePath);
            Dictionary<string, object> modelInfos = GetPropertyInfoArray(model);

            List<string> resultLines = new List<string>();

            bool isInGroup = false;
            List<string> groupLines = new List<string>();
            foreach (string line in tempLines)
            {
                if (line.Contains(KEY_GROUP_START))
                {
                    isInGroup = true;
                }
                else if (line.Contains(KEY_GROUP_END))
                {
                    // modelItem???
                    List<string> modeItems = AnalyzeTemplate(groupLines);
                    List<string> expandLines = new List<string>();

                    List<string> keysOfStringList = modeItems.FindAll(x => modelInfos[x] is List<string>);
                    List<string> keysOfString = modeItems.FindAll(x => modelInfos[x] is string);

                    // All objects with type of List<string> exist in the same Group should have same count of elements.
                    // Ex: format like {key} = {value}, count of key list should equals count of value list.
                    List<int> counts = new List<int>();
                    keysOfStringList.ForEach(x =>
                    {
                        counts.Add(((List<string>)modelInfos[x]).Count);
                    });

                    if (counts.Distinct().Count() != 1)
                    {
                        string msg = "All objects with type of List<string> exist in the same Group should have same count of elements. Please check these model properties: " +
                                     string.Join(",", keysOfStringList).Replace("{", "").Replace("}", "");
                        throw new FWFrameException(msg);
                    }

                    //Start to fill template with model
                    int repeatCount = ((List<string>)modelInfos[keysOfStringList.First()]).Count;
                    for (int i = 0; i < repeatCount; i++)
                    {
                        foreach (string groupLine in groupLines)
                        {
                            StringBuilder newLine = new StringBuilder(groupLine);
                            // fill item list into model
                            modeItems.ForEach(modeItem =>
                            {
                                if (keysOfStringList.Contains(modeItem))
                                {
                                    newLine.Replace(modeItem, ((List<string>)modelInfos[modeItem])[i]);
                                }
                                else
                                {
                                    newLine.Replace(modeItem, (string)modelInfos[modeItem]);
                                }
                            });
                            expandLines.Add(newLine.ToString());
                        }
                    }

                    resultLines.AddRange(expandLines);

                    isInGroup = false;
                    groupLines.Clear();
                }
                else
                {
                    if (isInGroup)
                    {
                        // add the lines in a group to a list of string and do the replacement when group end is reached.
                        groupLines.Add(line);
                    }
                    else
                    {
                        // if current line is not in a group, replace key words with value in modelInfos directly.
                        List<string> modeItems = AnalyzeTemplate(new List<string>() { line });
                        //modeItems.ForEach(x => line.Replace(x, modelInfos[x] as string));
                        //resultLines.Add(line.ToString());
                        if(modeItems != null)
                        {
                            StringBuilder lsb = new StringBuilder(line);
                            modeItems.ForEach(x => lsb.Replace(x, modelInfos[x] as string));
                            resultLines.Add(lsb.ToString());
                        }
                        else
                        {
                            resultLines.Add(line.ToString());
                        }
                    }
                }
            }

            return resultLines;
        }

        /// <summary>
        /// insert "{" and append "}" to each key of model
        /// </summary>
        /// <param name="model">The model.</param>
        /// <returns></returns>
        private static Dictionary<string, object> GetPropertyInfoArray(Dictionary<string, object> model)
        {
            Dictionary<string, object> modelInfos = new Dictionary<string, object>();

            foreach (var item in model)
            {
                modelInfos.Add("{" + item.Key + "}", item.Value);
            }
            return modelInfos;
        }

        /// <summary>
        /// look for characters in targetLines with format of "{*}" and return a distinct list of search result
        /// </summary>
        /// <param name="targetLines">The target lines.</param>
        /// <returns></returns>
        private static List<string> AnalyzeTemplate(List<string> targetLines)
        {
            List<string> modeItems = new List<string>();

            // 1. start with '{'
            // 2. end with '}'
            // 3. no '}' between '{' and '}'
            string pattern = @"\{[^}]*\}"; // \{ [^}]* \} ex: {XXXXX}
            Regex rgx = new Regex(pattern, RegexOptions.IgnoreCase);

            foreach (string line in targetLines)
            {
                Match m = rgx.Match(line);
                while (m.Success)
                {
                    foreach (Group group in m.Groups)
                    {
                        modeItems.Add(group.Value);
                    }

                    m = m.NextMatch();
                }
            }

            return modeItems.Distinct().ToList();
        }
    }
}
