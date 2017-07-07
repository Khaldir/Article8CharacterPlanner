using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace CharacterPlanner
{
    public partial class Form1 : Form
    {
        Microsoft.Office.Interop.Word.Application wordApp = null;
        Document wordDoc = null;
        string pMod1, pMod2, pMod3 = "";
        bool validAttributes = true;
        bool validAbilities = true;
        bool validMods = true;
        bool validBGs = true;
        bool validDefs = true;
        bool validFlaws = true;
        int[] AttributePoints = new int[3];
        int[] AbilityPoints = new int[3];
        string[] selectedBG = new string[4];
        CostedItem[] selectedFlaws = new CostedItem[4];
        int freebiePoints = 15;

        public Form1()
        {
            InitializeComponent();
            string[] bgs = new string[9] { "Allies", "Alternate Identity", "Contacts", "Fame", "Clearance", "Influence", "Mentor", "Resources", "Status" };
            CostedItem[] flaws = new CostedItem[24]
            {
                new CostedItem(1, "Hard Of Hearing"),
                new CostedItem(1, "Short"),
                new CostedItem(1, "Tic/Twitch"),
                new CostedItem(1, "Poor Sight"),
                new CostedItem(1, "Nightmares"),
                new CostedItem(1, "Shy"),
                new CostedItem(1, "Speech Impediment"),
                new CostedItem(1, "Amnesia"),
                new CostedItem(1, "Botched Presentation"),
                new CostedItem(1, "Dark Secret"),
                new CostedItem(1, "Expendable"),
                new CostedItem(1, "Mistaken Identity"),
                new CostedItem(1, "Sympathiser"),
                new CostedItem(2, "Disfigured"),
                new CostedItem(2, "Vengeful"),
                new CostedItem(2, "New Arrival"),
                new CostedItem(2, "Catspaw"),
                new CostedItem(2, "Old Flame"),
                new CostedItem(3, "Addiction"),
                new CostedItem(3, "Lazy"),
                new CostedItem(3, "Sleeping with the Enemy"),
                new CostedItem(1, "Enemy"),
                new CostedItem(2, "Enemy"),
                new CostedItem(3, "Enemy")
            };
            ComboBox[] backgroundBoxes = new ComboBox[4];
            backgroundBoxes[0] = bg1;
            backgroundBoxes[1] = bg2;
            backgroundBoxes[2] = bg3;
            backgroundBoxes[3] = bg4;
            for (int i = 0; i < 4; i++)
            {
                backgroundBoxes[i].Items.AddRange(bgs);
            }
            flawBox.Items.AddRange(flaws);
        }

        private void checkValid()
        {
            ExportButton.Enabled = validAbilities && validAttributes && validMods && validDefs && validBGs && validFlaws;
            if (!validAbilities)
                ExportButton.Text = "Invalid Abilities";
            else if (!validAttributes)
                ExportButton.Text = "Invalid Attributes";
            else if (!validMods)
                ExportButton.Text = "Invalid Prop. Mods";
            else if (!validDefs)
                ExportButton.Text = "Invalid Defenses";
            else if (!validBGs)
                ExportButton.Text = "Invalid Backgrounds";
            else if (!validFlaws)
                ExportButton.Text = "Invalid Flaws";
            else
                ExportButton.Text = "Export Character";
        }

        private void checkValidAttributes()
        {
            validAttributes = true;
            if (AttributePoints[0] == AttributePoints[1] || AttributePoints[0] == AttributePoints[2] || AttributePoints[1] == AttributePoints[2])
            {
                validAttributes = false;
            }
            validAttributes = Strength.Value + Dexterity.Value + Stamina.Value - 3 <= AttributePoints[0] && validAttributes;
            validAttributes = Charisma.Value + Manipulation.Value + Appearance.Value - 3 <= AttributePoints[1] && validAttributes;
            validAttributes = Perception.Value + Intelligence.Value + Wits.Value - 3 <= AttributePoints[2] && validAttributes;
            checkValid();
        }

        private void checkValidAbilities()
        {
            validAbilities = true;
            if (AbilityPoints[0] == AbilityPoints[1] || AbilityPoints[0] == AbilityPoints[2] || AbilityPoints[1] == AbilityPoints[2])
            {
                validAbilities = false;
            }
            validAbilities = Alertness.Value + Athletics.Value + Spook.Value + Brawl.Value + Empathy.Value + Expression.Value + Intimidate.Value + Leadership.Value + Streetwise.Value + Subterfuge.Value <= AbilityPoints[0] && validAbilities;
            validAbilities = Animal.Value + Crafts.Value + Drive.Value + Etiquette.Value + Firearms.Value + Larceny.Value + Melee.Value + Performance.Value + Pilot.Value + Stealth.Value + Survival.Value <= AbilityPoints[1] && validAbilities;
            validAbilities = Academics.Value + Computers.Value + Finance.Value + Investigation.Value + Law.Value + Medicine.Value + Netlore.Value + Politics.Value + Science.Value + Technology.Value + Demolitions.Value <= AbilityPoints[2] && validAbilities;
            checkValid();
        }

        private void AttributePointsChange(object sender, EventArgs e)
        {
            AttributePoints[0] = PhysicalPoints.SelectedIndex;
            AttributePoints[1] = SocialPoints.SelectedIndex;
            AttributePoints[2] = MentalPoints.SelectedIndex;

            for (int i = 0; i < 3; i++)
            {
                switch (AttributePoints[i])
                {
                    case 0: AttributePoints[i] = 3; break;
                    case 1: AttributePoints[i] = 5; break;
                    case 2: AttributePoints[i] = 7; break;
                }
            }
            checkValidAttributes();
        }

        private void AbilityPointsChange(object sender, EventArgs e)
        {
            AbilityPoints[0] = TalentPoints.SelectedIndex;
            AbilityPoints[1] = SkillPoints.SelectedIndex;
            AbilityPoints[2] = KnowPoints.SelectedIndex;
            
            for (int i = 0; i < 3; i++)
                {
                    switch (AbilityPoints[i])
                    {
                        case 0: AbilityPoints[i] = 6; break;
                        case 1: AbilityPoints[i] = 10; break;
                        case 2: AbilityPoints[i] = 14; break;
                    }
                }
            checkValidAbilities();
        }

        private void AttributeScroll(object sender, EventArgs e)
        {
            checkValidAttributes();
        }

        private void AbilityScroll(object sender, EventArgs e)
        {
            checkValidAbilities();
        }

        private void CorpSelected(object sender, EventArgs e)
        {
            string CorpDesc = "";
            switch (Corp.SelectedIndex)
            {
                //Mazad
                case 0:
                    {
                        CorpDesc = "Mazad Al-Zameel - A young up-and-comer in the megacorporate world. Secretive and bold, carrying on the tradition of the assassin \r\n\r\n Mazad Mods: Synth - Nerve Interface, Stealth Field, Enhanced Perception Suite";
                        pMod1 = "Synth-Nerve Interface";
                        pMod2 = "Stealth Field";
                        pMod3 = "Enhanced Perception Suite";
                        break;
                    }
                //Kirov
                case 1:
                    {
                        CorpDesc = "Kirov Heavy Works Foundation - A Russian heavy-hitter born from the ashes of the Soviet Union. Proud, loud and loaded for action \r\n\r\n Kirov Mods: Synth-Muscle, Subdermal Ballistic Mesh, S.D.I.R";
                        pMod1 = "Synth-Muscle";
                        pMod2 = "Subdermal Ballistic Mesh";
                        pMod3 = "S.D.I.R.";
                        break;
                    }
                //Regency
                case 2:
                    {
                        CorpDesc = "Regency Rejuvionics - A finishing school for the influential and consummate perfectionists of the human body. Beautiful, Alluring and potentially deadly. \r\n\r\n Regency Mods: S.D.I.R, Synth-Nerve Interface, Enhanced Perception Suite";
                        pMod1 = "S.D.I.R.";
                        pMod2 = "Synth-Nerve Interface";
                        pMod3 = "Enhanced Perception Suite";
                        break;
                    }
                //McAllister
                case 3:
                    {
                        CorpDesc = "McAllister Pharmaceuticals - A humanitarian medical organisation which hides a secret web of experimentation and torture. \r\n\r\n McAllister Mods: Enhanced Perception Suite, Dominance Suite, Necros Module";
                        pMod1 = "Enhanced Perception Suite";
                        pMod2 = "Dominance Suite";
                        pMod3 = "Necros Module";
                        break;
                    }
                //Novatek
                case 4:
                    {
                        CorpDesc = "Novatek - A all-rounder which strives to be everywhere, at all times. The very picture of the scummy corporate world, turned up to 11. \r\n\r\n Novatek Mods: Subdermal Ballistic Mesh, Dominance Suite, S.D.I.R";
                        pMod1 = "Subdermal Ballistic Mesh";
                        pMod2 = "Dominance Suite";
                        pMod3 = "S.D.I.R.";
                        break;
                    }
                //Multi-Arms Global
                case 5:
                    {
                        CorpDesc = "Multi-Arms Global - A mechanical corp bringing a utopian vision of automation to all aspects of life; war is just another avenue for business \r\n\r\n Multi-Arms Mods: Stealth Field, Enhanced Perception Suite, Drone Neural Network";
                        pMod1 = "Stealth Field";
                        pMod2 = "Enhanced Perception Suite";
                        pMod3 = "Drone Neural Network";
                        break;
                    }

            }

            propModParent.Enabled = true;
            prop1.Text = pMod1;
            prop2.Text = pMod2;
            prop3.Text = pMod3;
            CorpDescription.Text = CorpDesc;
        }

        private void ExportCharacter(object sender, EventArgs e)
        {
            //Initialise Word
            wordApp = new Microsoft.Office.Interop.Word.Application();
            wordApp.Visible = true;
            try
            {
                wordDoc = wordApp.Documents.Open(System.Windows.Forms.Application.StartupPath + "/blanksheet.docx");
            }
            catch
            {

            }


            //Character Details
            //Character Name
            getRange("CharacterName").Text = CharacterName.Text;

            //Nature
            getRange("Nature").Text = Nature.Text;

            //Demeanour
            getRange("Demeanour").Text = Demeanour.Text;

            //Corp
            getRange("CorpName").Text = Corp.Text;

            //Proprietary Mod Names - Inferred from Corp
            getRange("Proprietary1").Text = pMod1;
            getRange("Proprietary2").Text = pMod2;
            getRange("Proprietary3").Text = pMod3;

            //Attributes
            //Physical
            getRange("Strength").Text = getDots(Strength.Value);
            getRange("Dexterity").Text = getDots(Dexterity.Value);
            getRange("Stamina").Text = getDots(Stamina.Value);
            //Social
            getRange("Charisma").Text = getDots(Charisma.Value);
            getRange("Manipulation").Text = getDots(Manipulation.Value);
            getRange("Appearance").Text = getDots(Appearance.Value);
            //Mental
            getRange("Perception").Text = getDots(Perception.Value);
            getRange("Intelligence").Text = getDots(Intelligence.Value);
            getRange("Wits").Text = getDots(Wits.Value);
            //Inferred Values
            getRange("Willpower").Text = getDots(Stamina.Value + Math.Max(Intelligence.Value, Wits.Value), 10);

            //Abilities
            //Talents
            getRange("Alertness").Text = getDots(Alertness.Value);
            getRange("Athletics").Text = getDots(Athletics.Value);
            getRange("Spook").Text = getDots(Spook.Value);
            getRange("Brawl").Text = getDots(Brawl.Value);
            getRange("Empathy").Text = getDots(Empathy.Value);
            getRange("Expression").Text = getDots(Expression.Value);
            getRange("Intimidate").Text = getDots(Intimidate.Value);
            getRange("Leadership").Text = getDots(Leadership.Value);
            getRange("Streetwise").Text = getDots(Streetwise.Value);
            getRange("Subterfuge").Text = getDots(Subterfuge.Value);
            //Skills
            getRange("Animal").Text = getDots(Animal.Value);
            getRange("Crafts").Text = getDots(Crafts.Value);
            getRange("Drive").Text = getDots(Drive.Value);
            getRange("Etiquette").Text = getDots(Etiquette.Value);
            getRange("Firearms").Text = getDots(Firearms.Value);
            getRange("Larceny").Text = getDots(Larceny.Value);
            getRange("Melee").Text = getDots(Melee.Value);
            getRange("Performance").Text = getDots(Performance.Value);
            getRange("Pilot").Text = getDots(Pilot.Value);
            getRange("Stealth").Text = getDots(Stealth.Value);
            getRange("Survival").Text = getDots(Survival.Value);
            //Knowledges
            getRange("Academics").Text = getDots(Academics.Value);
            getRange("Computers").Text = getDots(Computers.Value);
            getRange("Finance").Text = getDots(Finance.Value);
            getRange("Investigation").Text = getDots(Investigation.Value);
            getRange("Law").Text = getDots(Law.Value);
            getRange("Medicine").Text = getDots(Medicine.Value);
            getRange("Netlore").Text = getDots(Netlore.Value);
            getRange("Politics").Text = getDots(Politics.Value);
            getRange("Science").Text = getDots(Science.Value);
            getRange("Technology").Text = getDots(Technology.Value);
            getRange("Demolitions").Text = getDots(Demolitions.Value);

            //Advantages
            //Proprietary Mods
            getRange("ProprietaryValue1").Text = getDots(propMod1.Value);
            getRange("ProprietaryValue2").Text = getDots(propMod2.Value);
            getRange("ProprietaryValue3").Text = getDots(propMod3.Value);
            //Backgrounds & Juice
            TrackBar[] bars = new TrackBar[4] { bgdots1, bgdots2, bgdots3, bgdots4 };
            bool dotsInClearance = false;
            for (int i = 1; i < 5; i++)
            {
                getRange("Background" + i.ToString()).Text = selectedBG[i-1];
                getRange("BackgroundDots" + i.ToString()).Text = getDots(bars[i - 1].Value);
                if(selectedBG[i-1]=="Clearance" && !dotsInClearance)
                {
                    dotsInClearance = true;
                    if(bars[i-1].Value < 4)
                        getRange("JuiceperTurn").Text = "1";
                    else if (bars[i - 1].Value == 4)
                        getRange("JuiceperTurn").Text = "2";
                    else if (bars[i - 1].Value == 5)
                        getRange("JuiceperTurn").Text = "3";


                    getRange("Juice").Text = getDots(10 + bars[i - 1].Value, 10 + bars[i - 1].Value);
                    getRange("ClearanceLevel").Text = (9 - bars[i - 1].Value).ToString();
                }
            }
            if (!dotsInClearance)
            {
                getRange("JuiceperTurn").Text = "1";
                getRange("Juice").Text = getDots(10, 10);
                getRange("ClearanceLevel").Text = (9).ToString();
            }
            //Cyber Defenses
            getRange("Firewall").Text = getDots(firewall.Value);
            getRange("Backtrace").Text = getDots(backtrace.Value);
            getRange("AttackBarrier").Text = getDots(attackBarrier.Value);

            //Freebie Points
            //Flaws
            for (int i = 1; i < 4; i++)
            {
                getRange("Flaw" + (i.ToString())).Text = selectedFlaws[i - 1].item;
                getRange("Flaw" + (i.ToString()) + "Pts").Text = selectedFlaws[i - 1].cost.ToString();
            }


            //Save Document
            wordDoc.SaveAs2(System.Windows.Forms.Application.StartupPath + "/" + CharacterName.Text + " - Character Sheet");
        }

        private void pModScroll(object sender, EventArgs e)
        {
            if(propMod1.Value+propMod2.Value+propMod3.Value > 3)
            {
                validMods = false;
            }
            else
            {
                validMods = true;
            }
            checkValid();
        }

        private void BackgroundChange(object sender, EventArgs e)
        {
            string[] bgs = new string[9] { "Allies", "Alternate Identity", "Contacts", "Fame", "Clearance", "Influence", "Mentor", "Resources", "Status" };
            ComboBox[] backgroundBoxes = new ComboBox[4];
            backgroundBoxes[0] = bg1;
            backgroundBoxes[1] = bg2;
            backgroundBoxes[2] = bg3;
            backgroundBoxes[3] = bg4;

            //Clear each box, then fill with fresh list
            for (int i = 0; i < 4; i++)
            {
                selectedBG[i] = (string)backgroundBoxes[i].SelectedItem;
                for (int n = 0; n < 9; n++)
                {
                    if (!backgroundBoxes[i].Items.Contains(bgs[n]))
                    {
                        backgroundBoxes[i].Items.Add(bgs[n]);
                    }
                }   
                for (int n = 0; n < 4; n++)
                {
                    //Remove already chosen options
                    if (n!=i && selectedBG[n]!="")
                    {
                        backgroundBoxes[i].Items.Remove(selectedBG[n]);
                    }
                }
            }

            


        }

        private Microsoft.Office.Interop.Word.Range getRange(string Tag)
        {
            Bookmark bkm = wordDoc.Bookmarks[Tag];
            return bkm.Range;
        }

        private void CyberDefensesScroll(object sender, EventArgs e)
        {
            if(firewall.Value+backtrace.Value+attackBarrier.Value-3 > 6)
            {
                validDefs = false;
            }
            checkValid();
        }

        private void flawsChanged(object sender, EventArgs e)
        {
            if (flawBox.CheckedItems.Count > 3)
            {
                validFlaws = false;
                checkValid();
            }
            else
            {
                freebiePoints = 15;
                foreach (CostedItem flaw in flawBox.CheckedItems)
                {
                    freebiePoints = freebiePoints + flaw.cost;
                }
            }
                
        }

        private string getDots(int dots)
        {
            string dotString = "";
            for (int i = 0; i < dots; i++)
            {
                dotString = dotString + "●";
            }
            for (int i = dots; i < 5; i++)
            {
                dotString = dotString + "○";
            }
            return dotString;
        }

        private string getDots(int dots, int maxDots)
        {
            string dotString = "";
            for (int i = 0; i < dots; i++)
            {
                dotString = dotString + "●";
            }
            for (int i = dots; i < maxDots; i++)
            {
                dotString = dotString + "○";
            }
            return dotString;
        }


    }

    public class CostedItem
    {
        public int cost;
        public string item;

        public CostedItem(int costIn, string itemin)
        {
            cost = costIn;
            item = itemin;
        }

        override public string ToString()
        {
            return (item + " - " + cost.ToString());
        }
    }
}
