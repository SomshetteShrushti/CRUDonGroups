using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.SharePoint.Client;

namespace CRUDOPerationsOnGroups
{
    public partial class Form1 : System.Windows.Forms.Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Method To get All users from the perticular group
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            using (ClientContext ctx = new ClientContext("URL"))
            {
                Web web = ctx.Web;
                GroupCollection groups = web.SiteGroups;
                Group group = groups.GetByName("HOD");
                UserCollection users = group.Users;
                ctx.Load(users);
                ctx.ExecuteQuery();
                foreach(User user in users)
                {
                    MessageBox.Show(user.Email+ "  " + user.LoginName);
                }
            }
        }

        /// <summary>
        /// Method To retrive specific properties of users from the groups
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            using (ClientContext ctx = new ClientContext("URL"))
            {
                Web web = ctx.Web;
                GroupCollection groups = web.SiteGroups;
                Group group = groups.GetByName("HOD");
                UserCollection users = group.Users;
                ctx.Load(users, pusers => pusers.Include(user => user.Title, user => user.LoginName, user => user.Email));
                ctx.ExecuteQuery();
                foreach (User ouser in users)
                {
                    MessageBox.Show(ouser.Title + " " +ouser.LoginName + " " + ouser.Email);

                }
            }
        }

        /// <summary>
        /// Method to retrive all the users from all the groups present on the site collection
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            using (ClientContext ctx = new ClientContext("URL"))
            {
                Web web = ctx.Web;
                GroupCollection collgroups = web.SiteGroups;
                ctx.Load(collgroups);

                ctx.Load(collgroups,
                    groups => groups.Include(
                        group => group.Users));
                ctx.ExecuteQuery();
                foreach (Group oGroup in collgroups)
                {
                    UserCollection collUser = oGroup.Users;

                    foreach (User oUser in collUser)
                    {
                        MessageBox.Show(oGroup.Id+" "+oGroup.Title + " " + oUser.Title + " " + oUser.LoginName);
                    }
                }
            }
        }
        /// <summary>
        /// Method to add USER in the perticular list 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            using (ClientContext ctx = new ClientContext("URL"))
            {
                Web web = ctx.Web;
                GroupCollection groups = web.SiteGroups;
                Group group = groups.GetByName("HOD");
                UserCreationInformation userCreationInformation = new UserCreationInformation();
                userCreationInformation.LoginName = @"domain\hrmsemployee";
                userCreationInformation.Email = "ssr@gmail.com";
                User user = group.Users.Add(userCreationInformation);
                ctx.ExecuteQuery();
            }
            MessageBox.Show("User Added to the group");
        }
        /// <summary>
        /// Method to update the information of user present in the group
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button5_Click(object sender, EventArgs e)
        {
            using (ClientContext ctx = new ClientContext("URL"))
            {
                Web web = ctx.Web;
                GroupCollection groups = web.SiteGroups;
                Group group = groups.GetByName("HOD");
                UserCollection users = group.Users;
                User user = users.GetByLoginName(@"Domain\hrmsemployee");
                ctx.Load(user);
                user.Email = "rudra@gmail.com";
                user.Update();
                ctx.ExecuteQuery();
            }
            MessageBox.Show("User Updated Sucessfully");
        }

        /// <summary>
        /// Method to Delete User from the Group
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button6_Click(object sender, EventArgs e)
        {
            using (ClientContext ctx = new ClientContext("http://hr.zubaircorp.com/"))
            {
                Web web = ctx.Web;
                GroupCollection groups = web.SiteGroups;
                Group group = groups.GetByName("HOD");
                UserCollection users = group.Users;
                group.Users.RemoveByLoginName(@"zubaircorp\hrmsemployee");
                ctx.ExecuteQuery();
            }
            MessageBox.Show("User Removed Sucessfully");
        }
        /// <summary>
        /// Method to create new group in sharepoint site
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button7_Click(object sender, EventArgs e)
        {
            using (ClientContext ctx = new ClientContext("http://hr.zubaircorp.com/"))
            {
                Web web = ctx.Web;
                GroupCreationInformation groupCreation = new GroupCreationInformation();
                groupCreation.Title = "Testing123";
                Group group = web.SiteGroups.Add(groupCreation);
                ctx.ExecuteQuery();
            }
            MessageBox.Show("Group Created Successfully");
        }
        /// <summary>
        /// Method to Update the group name 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button8_Click(object sender, EventArgs e)
        {
            using (ClientContext ctx = new ClientContext("http://hr.zubaircorp.com/"))
            {
                Web web = ctx.Web;
                Group group = web.SiteGroups.GetByName("Testing123");
                group.Title = "Testing34";
                group.Update();
                ctx.ExecuteQuery();
            }
            MessageBox.Show("Group Name UPdated Successfully");
        }
        /// <summary>
        /// Remove/ Delete group from site
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button9_Click(object sender, EventArgs e)
        {
            using (ClientContext ctx = new ClientContext("http://hr.zubaircorp.com/"))
            {
                Web web = ctx.Web;
                Group group = web.SiteGroups.GetByName("Testing34");
                web.SiteGroups.Remove(group);
                ctx.ExecuteQuery();
            }
            MessageBox.Show("Group Deleted Successfully");
        }
    }
}
