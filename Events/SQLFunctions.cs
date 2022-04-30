using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Events
{
    public class SQLfunctions
    {
        string adress = @"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=EventsBD;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False";

        public bool DeleteFromGridByIndex(string Table, int id)
        {
            SqlConnection con = new SqlConnection(adress);
            SqlCommand cmd = new SqlCommand();

            con.Open();
            cmd.CommandText = "Delete from " + Table + " where Id = " + id ;
            cmd.Connection = con;
            
            try
            {
                cmd.ExecuteNonQuery();
            }
            catch (SqlException ex)
            {
                return false;
            }
            con.Close();
            return true;
        }


        #region Comboboxes
        public DataTable SelectInComboboxClient()
        {
            DataTable dt = new DataTable();
            SqlConnection con = new SqlConnection(adress);
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter adapt = new SqlDataAdapter();
            con.Open();
            cmd.CommandText = "Select Id as Clientid,ConCat(SurName,' ',Name,' ',LastName) as FIO from Clients";
            cmd.Connection = con;
            adapt.SelectCommand = cmd;
            adapt.Fill(dt);
            con.Close();
            return dt;
        }
        public DataTable SelectInComboboxResponsible()
        {
            DataTable dt = new DataTable();
            SqlConnection con = new SqlConnection(adress);
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter adapt = new SqlDataAdapter();
            con.Open();
            cmd.CommandText = "Select id,ConCat(SurName,' ',Name,' ',LastName) as FIO from Responsible";
            cmd.Connection = con;
            adapt.SelectCommand = cmd;
            adapt.Fill(dt);
            con.Close();
            return dt;
        }
        public DataTable SelectInComboboxAdress(int TypeId)
        {
            DataTable dt = new DataTable();
            SqlConnection con = new SqlConnection(adress);
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter adapt = new SqlDataAdapter();
            con.Open();
            cmd.CommandText = @"Select id,ConCat('|',Name,'| ',Adress) as Adress from Adresses where Adresses.id_Type = " + TypeId;
            cmd.Connection = con;
            adapt.SelectCommand = cmd;
            adapt.Fill(dt);
            con.Close();
            return dt;
        }
        public DataTable SelectInComboboxType()
        {
            DataTable dt = new DataTable();
            SqlConnection con = new SqlConnection(adress);
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter adapt = new SqlDataAdapter();
            con.Open();
            cmd.CommandText = "Select id,Type from TypesOfEvent";
            cmd.Connection = con;
            adapt.SelectCommand = cmd;
            adapt.Fill(dt);
            con.Close();
            return dt;
        }
        #endregion
        #region EventGrid
        public void AddEvent(string Name, int IdClient, int IdRespons, int IdAdress, string addInf, string dateTime)
        {
            SqlConnection con = new SqlConnection(adress);
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = "Insert into Events(Name, id_client,id_responsible,id_adress,add_inf,dateTime) Values(N'" + Name + "',N'" + IdClient + "',N'" + IdRespons + "', N'" + IdAdress + "', N'" + addInf + "',N'" + dateTime + "')";
            cmd.Connection = con;
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
        }
        public void EditEvent(string Name, int IdClient, int IdRespons, int IdAdress, string addInf, string dateTime,int id)
        {
            SqlConnection con = new SqlConnection(adress);
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = "Update Events set Name = N'" + Name + "', id_client = N'" + IdClient + "' , id_responsible = N'" + IdRespons + "' , id_adress = N'" + IdAdress + "' , add_inf = N'" + addInf + "', dateTime = N'" + dateTime + "' where Id = "+ id;
            cmd.Connection = con;
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
        }
        
        public DataTable SelectAllFromEvents()
        {
            DataTable dt = new DataTable();
            SqlConnection con = new SqlConnection(adress);
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter adapt = new SqlDataAdapter();

            con.Open();
            cmd.CommandText = "Select Events.Id,Events.Name as Название, ConCat(Clients.SurName,' ',Clients.Name,' ', Clients.LastName) as Клиент, ConCat(Responsible.SurName,' ' , Responsible.Name,' ', Responsible.LastName) as Ответственное_лицо, ConCat(TypesOfEvent.Type,' |',Adresses.Name,'| ',Adresses.Adress  ) as Помещение, Events.add_inf as ДопИнформация, Events.dateTime From Events Left Join Adresses On Adresses.Id = Events.id_adress inner join TypesOfEvent on Adresses.id_Type = TypesOfEvent.Id Left Join Clients On Clients.Id = Events.id_client Left Join Responsible On Responsible.Id = Events.id_responsible";
            cmd.Connection = con;
            adapt.SelectCommand = cmd;
            adapt.Fill(dt);
            con.Close();
            return dt;
        }
        #endregion
        #region Clients
        public DataTable SelectAllClients_Responsible(string Table)
        {
            DataTable dt = new DataTable();
            SqlConnection con = new SqlConnection(adress);
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter adapt = new SqlDataAdapter();
            con.Open();
            cmd.CommandText = "Select * from " + Table;
            cmd.Connection = con;
            adapt.SelectCommand = cmd;
            adapt.Fill(dt);
            con.Close();
            return dt;
        }
        public void AddClient_responsible(string Table, string SurName, string Name,string LastName, string Phone)
        {
            SqlConnection con = new SqlConnection(adress);
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = "Insert into "+Table+"(SurName,Name,LastName,Phone) Values(N'" + SurName + "',N'" + Name + "',N'" + LastName + "', N'" + Phone + "')";
            cmd.Connection = con;
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
        }
        public void EditClientResponsible(string Table, string SurName, string Name, string LastName, string Phone, int Id)
        {
            SqlConnection con = new SqlConnection(adress);
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = "Update "+ Table + " set SurName = N'" + SurName + "', Name = N'" + Name + "' , LastName = N'" + LastName + "' , Phone = N'" + Phone + "'  where Id =" + Id;
            cmd.Connection = con;
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
        }
        #endregion

        #region AdressPan
        public DataTable SelectAllFromAdresses()
        {
            DataTable dt = new DataTable();
            SqlConnection con = new SqlConnection(adress);
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter adapt = new SqlDataAdapter();
            con.Open();
            cmd.CommandText = "Select Adresses.Id , TypesOfEvent.Type, Adresses.Name, Adresses.Adress from Adresses left join TypesOfEvent On TypesOfEvent.Id = Adresses.id_Type";
            cmd.Connection = con;
            adapt.SelectCommand = cmd;
            adapt.Fill(dt);
            con.Close();
            return dt;
        }
        public void EditAdress(string Name, string Adress, int TypeId, int id)
        {
            SqlConnection con = new SqlConnection(adress);
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = "Update Adresses set Name = N'" + Name + "', Adress = N'" + Adress + "' , id_Type = N'" + TypeId + "' where Id = "+  id;
            cmd.Connection = con;
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
        }
        public void AddAdress(string Name, string Adress, int TypeId)
        {
            SqlConnection con = new SqlConnection(adress);
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = "Insert into Adresses(Name, Adress,id_Type) Values(N'" + Name + "',N'" + Adress + "',N'" + TypeId + "')";
            cmd.Connection = con;
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
        }
        #endregion
        public DataTable Take_all_records_by_date(DateTimePicker firstDate, DateTimePicker lastDate)
        {
            SqlConnection con = new SqlConnection(); ;
            SqlCommand cmd = new SqlCommand(); ;
            DataTable table = new DataTable();   //объявляю адаптер, таблицу, конектор, команду
            SqlDataAdapter adapt = new SqlDataAdapter();
            con.ConnectionString = adress;    //Присваиваю коннектору строку подключения к бд  SELECT T1.key, T2.name FROM table1 T1 JOIN table2 T2 ON T1.name = T2.key;
            cmd.Connection = con;                       //комманде задается коннектор
            con.Open();
            cmd.CommandText = "Select Events.Id,Events.Name as Название, ConCat(Clients.SurName,' ',Clients.Name,' ', Clients.LastName) as Клиент, ConCat(Responsible.SurName,' ' , Responsible.Name,' ', Responsible.LastName) as Ответственное_лицо, ConCat(TypesOfEvent.Type,' |',Adresses.Name,'| ',Adresses.Adress  ) as Помещение, Events.add_inf as ДопИнформация, Events.dateTime From Events Left Join Adresses On Adresses.Id = Events.id_adress inner join TypesOfEvent on Adresses.id_Type = TypesOfEvent.Id Left Join Clients On Clients.Id = Events.id_client Left Join Responsible On Responsible.Id = Events.id_responsible where Events.dateTime between '" + firstDate.Value.ToString("yyyy-MM-dd") + "' and '" + lastDate.Value.ToString("yyyy-MM-dd") + "'";
            adapt.SelectCommand = cmd;
            adapt.Fill(table);   //заполняю таблицу полученными данными
            return table;
        }
    }
}
