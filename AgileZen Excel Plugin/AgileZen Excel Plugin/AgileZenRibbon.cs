using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using System.Xml;


namespace AgileZen_Excel_Plugin
{
    
    public partial class AgileZenRibbon
    {
        string _apiKey = "?apikey=e7907dab6ceb4dbab37261515c482b75";
        string _baseUrl = "https://agilezen.com/api/v1/";
        
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            
        }

        private void projectDropDown_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            RibbonFactory ribbonFactory = Globals.Factory.GetRibbonFactory();
            RibbonDropDownItem dditem;

            string projectId = (string)ProjectsDropDown.SelectedItem.Tag;
            XmlNodeList phases = getAgileZenDoc("projects/"+ projectId +"/phases.xml").SelectNodes("//phase");
          
            foreach ( XmlNode phase in phases){
                        dditem =  ribbonFactory.CreateRibbonDropDownItem();
                        dditem.Label = phase.SelectSingleNode("name").InnerText; 
                        dditem.Tag =  phase.SelectSingleNode("id").InnerText;
                        PhaseDropDown.Items.Add(dditem);
            }
        }

        private void PhaseDropDown_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            string projectId = (string)ProjectsDropDown.SelectedItem.Tag;
            string phaseId = (string)PhaseDropDown.SelectedItem.Tag;
            XmlNodeList stories = getAgileZenDoc("projects/" + projectId + "/phases./"+
                phaseId+"/stories.xml").SelectNodes("//story");

            List<AgileZenStory> storyList = new List<AgileZenStory>();
            AgileZenStory azStory;

            foreach (XmlNode story in stories)
            {
                azStory = new AgileZenStory();
                azStory.id = story.SelectSingleNode("id").InnerText;
                azStory.color = story.SelectSingleNode("color").InnerText;
                azStory.phase = story.SelectSingleNode("phase/name").InnerText;
                azStory.text = story.SelectSingleNode("text").InnerText;
                 
                storyList.Add(azStory);                
            }

            writeStoriesToSheet(storyList);

        }

        private void writeStoriesToSheet(List<AgileZenStory> storyList)
        {
            Worksheet activeWorksheet = ((Worksheet) Globals.ThisAddIn.Application.ActiveSheet);
            Range range = activeWorksheet.get_Range("a1");
            foreach ( AgileZenStory  story in storyList) 
            {
                foreach ( System.Reflection.PropertyInfo info in story.GetType().GetProperties() ) 
                {
                    range.Value = info.GetValue(story, null);
                    range = range.get_Offset(0, 1);
                }
                range = range.get_Offset( 1,-story.GetType().GetProperties().Count());
                
            }
            
        }

        private void LogonButton_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonFactory ribbonFactory = Globals.Factory.GetRibbonFactory();
            RibbonDropDownItem dditem;

            XmlNodeList projects = getAgileZenDoc("projects.xml").SelectNodes("//project");
  
            foreach ( XmlNode project in projects)
            {
                dditem =  ribbonFactory.CreateRibbonDropDownItem();
                dditem.Label = project.SelectSingleNode("name").InnerText; 
                dditem.Tag =  project.SelectSingleNode("id").InnerText;
                ProjectsDropDown.Items.Add(dditem);
            }
        }

        private XmlDocument getAgileZenDoc(string prameterUrl)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(_baseUrl + prameterUrl + _apiKey +
                "&pageSize=1000");
            return doc  ;
        }
    }
}
