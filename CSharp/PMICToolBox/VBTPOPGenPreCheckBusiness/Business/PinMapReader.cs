using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace VBTPOPGenPreCheckBusiness.Business
{
    public class PinMapReader
    {
        public PinMapSheet ReadSheet(string pinMapFilePath)
        {
            PinMapSheet pinMapSheet = new PinMapSheet();
            int index = 0;
            int groupNameIndex = -1;
            int pinNameIndex = -1;
            int typeIndex = -1;
            string pinName, groupName, type;
            using (StreamReader sr = new StreamReader(pinMapFilePath))
            {
                while (!sr.EndOfStream)
                {
                    string line = sr.ReadLine();
                    index++;
                    if (line != null)
                    {
                        string[] arr = line.Split(new[] { '\t' }, StringSplitOptions.None);
                        if(groupNameIndex > 0)
                        {
                            if(arr.Length >= typeIndex + 1)
                            {
                                groupName = arr[groupNameIndex].Trim();
                                pinName = arr[pinNameIndex].Trim();
                                type = arr[typeIndex].Trim();
                                if (groupName == "" && pinName != "")
                                {
                                    pinMapSheet.AddPin(pinName, type);
                                }
                                else if (groupName != "")
                                {
                                    PinGroup existedGroup = pinMapSheet.GetPinGroupByName(groupName);
                                    if (existedGroup == null)
                                    {
                                        pinMapSheet.AddPinGroup(groupName,type);
                                        existedGroup = pinMapSheet.GetPinGroupByName(groupName);
                                    }
                                    existedGroup.AddPin(pinName);
                                }
                                else
                                {
                                    break;
                                }
                            }
                        }else
                        {
                            if(line.Contains("Group Name"))
                            {
                                groupNameIndex = arr.ToList().IndexOf("Group Name");
                                pinNameIndex = groupNameIndex + 1;
                                typeIndex = groupNameIndex + 2;
                            }
                        }

                       
                    }
                }
            }        
            return pinMapSheet;
        }
    }

    public class PinMapSheet
    {
        public List<Pin> PinList;
        public List<PinGroup> PinGroupList;
        public PinMapSheet()
        {
            PinList = new List<Pin>();
            PinGroupList = new List<PinGroup>();
        }
        public void AddPin(string pinName, string pinType)
        {
            if (!PinList.Exists(s => s.pinName.Equals(pinName, StringComparison.OrdinalIgnoreCase)))
                PinList.Add(new Pin(pinName,pinType));
        }

        public PinGroup GetPinGroupByName(string groupName)
        {
            return PinGroupList.Find(s => s.GroupName.Equals(groupName, StringComparison.OrdinalIgnoreCase));
        }

        public void AddPinGroup(string groupName, string pinType)
        {
            if (!PinGroupList.Exists(s => s.GroupName.Equals(groupName, StringComparison.OrdinalIgnoreCase)))
                PinGroupList.Add(new PinGroup(groupName, pinType));
        }
    }

    public class Pin
    {
        public string pinName;
        public string pinType;
        public Pin(string pinName, string pinType)
        {
            this.pinName = pinName;
            this.pinType = pinType;
        }
    }

    public class PinGroup
    {
        public string GroupName;
        public string PinType;
        public List<string> PinList;
        public PinGroup(string groupName, string pinType)
        {
            this.GroupName = groupName;
            this.PinType = pinType;
            PinList = new List<string>();
        }

        public void AddPin(string pinName)
        {
            if(!PinList.Exists(s=>s.Equals(pinName,StringComparison.OrdinalIgnoreCase)))
                PinList.Add(pinName);
        }
    }
   
    
}