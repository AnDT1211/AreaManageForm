using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AreaAddForm
{
    public partial class Form1 : Form
    {
        SldWorks swApp;

        ModelDoc2 swModel;

        SelectionMgr swSelMgr;

        Face2 swFace = default(Face2);

        Configuration config;
        CustomPropertyManager cusPropMgr;

        int lRetVal;
        object[] vPropNames = null;
        object vPropNamesObject;
        object vPropTypes;
        object vPropValues;
        string ValOut;
        string ResolvedValOut;
        bool wasResolved;
        object resolved;
        int nNbrProps;

        int custPropType;
        string[] listItem;
        public void khoitaoComboBox()
        {
            listItem = new string[] { "Loại sơn 1", "Loại sơn 2", "Loại sơn 3", "Loại sơn 4" };
            foreach (var item in listItem)
            {
                comboBox1.Items.Add(item);
            }
            comboBox1.SelectedItem = listItem[0];
            
        }

        public void setDefProp()
        {
            // Khoi tao Properties
            swApp = GetSW.GetApplication();

            swModel = (ModelDoc2)swApp.ActiveDoc;
            config = (Configuration)swModel.GetActiveConfiguration();

            cusPropMgr = config.CustomPropertyManager;

            nNbrProps = cusPropMgr.Count;

            // Get the names of the custom properties
            lRetVal = cusPropMgr.GetAll2(ref vPropNamesObject, ref vPropTypes, ref vPropValues, ref resolved);
            vPropNames = (object[])vPropNamesObject;

            int flag = 0;
            for (int i = 0; i < listItem.Length; i++)
            {
                for (int j = 0; j <= nNbrProps - 1; j++)
                {
                    custPropType = cusPropMgr.GetType2((string)vPropNames[j]);
                    string nana = Convert.ToString(vPropNames[j]);

                    if (nana == listItem[i])
                    {
                        flag = 1;
                        break;
                    }
                }
                if (flag == 0)
                {
                    cusPropMgr.Add3(listItem[i], (int)swCustomInfoType_e.swCustomInfoDouble, "0", (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                    flag = 0;
                }
                flag = 0;
            }
        }

        public Form1()
        {
            InitializeComponent();

            khoitaoComboBox();
            setDefProp();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            swApp = GetSW.GetApplication();

            swModel = swApp.ActiveDoc as ModelDoc2;

            swSelMgr = swModel.SelectionManager as SelectionMgr;
            
            swModel = (ModelDoc2)swApp.ActiveDoc;
            config = (Configuration)swModel.GetActiveConfiguration();

            cusPropMgr = config.CustomPropertyManager;

            //          Combo Box           //
            var item = this.comboBox1.GetItemText(this.comboBox1.SelectedItem);
            //          Combo Box           //

            string nameProp = Convert.ToString(item);


            cusPropMgr.Get5(nameProp, false, out ValOut, out ResolvedValOut, out wasResolved);

            double dienTich = Convert.ToDouble(ResolvedValOut);

            int tongCacMat = 0;
            for (int i = 0; i < swSelMgr.GetSelectedObjectCount2(-1); i++)
            {
                tongCacMat++;
                swFace = (Face2)swSelMgr.GetSelectedObject6(i + 1, -1);
                double x = (double)swFace.GetArea() * 1000000;
                dienTich += x;
            }
            cusPropMgr.Set2(nameProp, Convert.ToString(dienTich));
            MessageBox.Show("Đã cộng diện tích của " + Convert.ToString(tongCacMat) + " mặt vào " + nameProp);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            swApp = GetSW.GetApplication();

            swModel = swApp.ActiveDoc as ModelDoc2;

            swSelMgr = swModel.SelectionManager as SelectionMgr;


            swModel = (ModelDoc2)swApp.ActiveDoc;
            config = (Configuration)swModel.GetActiveConfiguration();

            cusPropMgr = config.CustomPropertyManager;

            //          Combo Box           //
            var item = this.comboBox1.GetItemText(this.comboBox1.SelectedItem);
            //          Combo Box           //

            string nameProp = Convert.ToString(item);


            cusPropMgr.Get5(nameProp, false, out ValOut, out ResolvedValOut, out wasResolved);

            double dienTich = Convert.ToDouble(ResolvedValOut);

            int tongCacMat = 0;
            for (int i = 0; i < swSelMgr.GetSelectedObjectCount2(-1); i++)
            {
                tongCacMat++;
                swFace = (Face2)swSelMgr.GetSelectedObject6(i + 1, -1);
                double x = (double)swFace.GetArea() * 1000000;
                dienTich -= x;
            }
            if (dienTich < 0)
            {
                dienTich = 0;
            }
            cusPropMgr.Set2(nameProp, Convert.ToString(dienTich));
            MessageBox.Show("Đã trừ diện tích của " + Convert.ToString(tongCacMat) + " mặt vào " + nameProp);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            setDefProp();
        }
    }
}
