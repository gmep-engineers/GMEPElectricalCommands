using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Xml.Linq;
using ElectricalCommands.Equipment;
using MySql.Data.MySqlClient;

namespace GMEPElectricalCommands.GmepDatabase
{
  public class GmepDatabase
  {
    public string ConnectionString { get; set; }
    public static MySqlConnection Connection { get; set; }

    public GmepDatabase()
    {
      ConnectionString = Properties.Settings.Default.ConnectionString;
      Connection = new MySqlConnection(ConnectionString);
    }

    public void OpenConnection()
    {
      if (Connection.State == System.Data.ConnectionState.Closed)
      {
        Connection.Open();
      }
    }

    public void CloseConnection()
    {
      if (Connection.State == System.Data.ConnectionState.Open)
      {
        Connection.Close();
      }
    }

    public List<Feeder> GetFeeders(string projectId)
    {
      List<Feeder> feeders = new List<Feeder>();
      //string query =
      //  @"
      //  SELECT * FROM electrical_panels
      //  WHERE electrical_panels.project_id = @projectId
      //  UNION
      //  SELECT * FROM electrical_transformers
      //  WHERE electrical_transformers.project_id = @projectId";
      string query =
        @"
        SELECT * FROM electrical_panels
        LEFT JOIN electrical_panel_bus_amp_ratings
        ON electrical_panel_bus_amp_ratings.id = electrical_panels.bus_amp_rating_id
        WHERE electrical_panels.project_id = @projectId";
      OpenConnection();
      MySqlCommand command = new MySqlCommand(query, Connection);
      command.Parameters.AddWithValue("projectId", projectId);
      MySqlDataReader reader = command.ExecuteReader();
      while (reader.Read())
      {
        feeders.Add(
          new Feeder(
            reader.GetString("id"),
            reader.GetString("name"),
            reader.GetString("parent_id"),
            "Panel",
            reader.GetInt32("parent_distance"),
            reader.GetFloat("loc_x"),
            reader.GetFloat("loc_y")
          )
        );
      }
      reader.Close();
      return feeders;
    }

    public List<Equipment> GetEquipment(string projectId)
    {
      List<Equipment> equip = new List<Equipment>();
      string query =
        @"
        SELECT
        electrical_equipment.id,
        electrical_equipment.parent_id,
        electrical_equipment.equip_no,
        electrical_equipment_categories.category,
        electrical_equipment.description,
        electrical_equipment_voltages.voltage,
        electrical_equipment.is_three_phase,
        electrical_equipment.parent_distance,
        electrical_equipment.loc_x,
        electrical_equipment.loc_y
        FROM electrical_equipment
        LEFT JOIN electrical_panels
        ON electrical_panels.id = electrical_equipment.parent_id
        LEFT JOIN electrical_equipment_categories
        ON electrical_equipment.category_id = electrical_equipment_categories.id
        LEFT JOIN electrical_equipment_voltages
        ON electrical_equipment_voltages.id = electrical_equipment.voltage_id
        WHERE electrical_equipment.project_id = @projectId";
      this.OpenConnection();
      MySqlCommand command = new MySqlCommand(query, Connection);
      command.Parameters.AddWithValue("projectId", projectId);
      MySqlDataReader reader = command.ExecuteReader();
      while (reader.Read())
      {
        string feederId = reader.GetString("parent_id");
        bool is3Phase = reader.GetInt32("is_three_phase") == 1;
        if (!String.IsNullOrEmpty(feederId))
        {
          equip.Add(
            new Equipment(
              reader.GetString("id"),
              reader.GetString("parent_id"),
              reader.GetString("equip_no"),
              "",
              reader.GetString("category"),
              reader.GetString("description"),
              reader.GetInt32("voltage"),
              is3Phase,
              reader.GetInt32("parent_distance"),
              reader.GetFloat("loc_x"),
              reader.GetFloat("loc_y")
            )
          );
        }
      }
      reader.Close();
      return equip;
    }

    public string GetProjectId(string projectNo)
    {
      Console.WriteLine(projectNo);
      string query = @"SELECT id FROM projects WHERE gmep_project_no = @projectNo";
      OpenConnection();
      MySqlCommand command = new MySqlCommand(query, Connection);
      command.Parameters.AddWithValue("@projectNo", projectNo);
      MySqlDataReader reader = command.ExecuteReader();
      string id = "";
      if (reader.Read())
      {
        id = reader.GetString("id");
      }
      reader.Close();
      return id;
    }
  }
}
