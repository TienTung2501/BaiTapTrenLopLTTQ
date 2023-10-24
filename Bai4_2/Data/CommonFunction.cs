using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
namespace Bai4_2.Data
{
    internal class CommonFunction
    {
        public void fillComboBox(ComboBox comboBoxName, DataTable data,string displayMember,string valueMember)
        {
            comboBoxName.DataSource = data;
            comboBoxName.DisplayMember = displayMember;
            comboBoxName.ValueMember = valueMember;

        }
    }
}
