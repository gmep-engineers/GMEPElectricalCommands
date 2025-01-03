using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Office2010.CustomUI;
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

    public List<Service> GetServices(string projectId)
    {
      List<Service> services = new List<Service>();
      string query =
        @"SELECT 
          electrical_services.id,
          electrical_services.name,
          electrical_service_meter_configs.meter_config,
          electrical_service_amp_ratings.amp_rating,
          electrical_service_voltages.voltage
          FROM `electrical_services`
          LEFT JOIN electrical_service_meter_configs
          ON electrical_services.electrical_service_meter_config_id = electrical_service_meter_configs.id
          LEFT JOIN electrical_service_amp_ratings
          ON electrical_service_amp_ratings.id = electrical_services.electrical_service_amp_rating_id
          LEFT JOIN electrical_service_voltages
          ON electrical_service_voltages.id = electrical_services.electrical_service_voltage_id
          WHERE electrical_services.project_id = @projectId";
      OpenConnection();
      MySqlCommand command = new MySqlCommand(query, Connection);
      command.Parameters.AddWithValue("projectId", projectId);
      MySqlDataReader reader = command.ExecuteReader();
      while (reader.Read())
      {
        services.Add(
          new Service(
            reader.GetString("id"),
            reader.GetString("name"),
            reader.GetString("meter_config"),
            reader.GetInt32("amp_rating"),
            reader.GetString("voltage")
          )
        );
      }
      reader.Close();
      return services;
    }

    public List<Panel> GetPanels(string projectId)
    {
      List<Panel> panels = new List<Panel>();
      string query =
        @"
        SELECT
        electrical_panels.id,
        electrical_panels.parent_id,
        electrical_panels.name,
        electrical_panels.parent_distance,
        electrical_panels.loc_x,
        electrical_panels.loc_y,
        electrical_panels.is_distribution,
        electrical_panel_bus_amp_ratings.amp_rating,
        electrical_service_voltages.voltage
        FROM electrical_panels
        LEFT JOIN electrical_panel_bus_amp_ratings
        ON electrical_panel_bus_amp_ratings.id = electrical_panels.bus_amp_rating_id
        LEFT JOIN electrical_service_voltages
        ON electrical_service_voltages.id = electrical_panels.voltage_id
        WHERE electrical_panels.project_id = @projectId";
      OpenConnection();
      MySqlCommand command = new MySqlCommand(query, Connection);
      command.Parameters.AddWithValue("projectId", projectId);
      MySqlDataReader reader = command.ExecuteReader();
      while (reader.Read())
      {
        panels.Add(
          new Panel(
            reader.GetString("id"),
            reader.GetString("parent_id"),
            reader.GetString("name"),
            reader.GetInt32("parent_distance"),
            reader.GetFloat("loc_x"),
            reader.GetFloat("loc_y"),
            reader.GetInt32("is_distribution"),
            0,
            reader.GetInt32("amp_rating"),
            reader.GetString("voltage")
          )
        );
      }
      reader.Close();
      return panels;
    }

    public List<Transformer> GetTransformers(string projectId)
    {
      List<Transformer> xfmrs = new List<Transformer>();
      string query =
        @"
        SELECT * FROM electrical_transformers
        LEFT JOIN electrical_transformer_kva_ratings
        ON electrical_transformer_kva_ratings.id = electrical_transformers.kva_id
        LEFT JOIN electrical_transformer_voltages
        ON electrical_transformer_voltages.id = electrical_transformers.voltage_id
        WHERE electrical_transformers.project_id = @projectId";
      OpenConnection();
      MySqlCommand command = new MySqlCommand(query, Connection);
      command.Parameters.AddWithValue("projectId", projectId);
      MySqlDataReader reader = command.ExecuteReader();
      while (reader.Read())
      {
        xfmrs.Add(
          new Transformer(
            reader.GetString("id"),
            reader.GetString("parent_id"),
            reader.GetString("name"),
            reader.GetInt32("parent_distance"),
            reader.GetFloat("loc_x"),
            reader.GetFloat("loc_y"),
            reader.GetFloat("kva_rating"),
            reader.GetString("voltage")
          )
        );
      }
      reader.Close();
      return xfmrs;
    }

    public List<Equipment> GetEquipment(string projectId)
    {
      List<Equipment> equip = new List<Equipment>();
      string query =
        @"
        SELECT
        electrical_equipment.id,
        electrical_equipment.parent_id,
        electrical_panels.name,
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
        bool is3Phase = reader.GetInt32("is_three_phase") == 1;
        equip.Add(
          new Equipment(
            reader.GetString("id"),
            reader.GetString("parent_id"),
            reader.GetString("name"),
            reader.GetString("equip_no"),
            reader.GetString("description"),
            reader.GetString("category"),
            reader.GetInt32("voltage"),
            is3Phase,
            reader.GetInt32("parent_distance"),
            reader.GetFloat("loc_x"),
            reader.GetFloat("loc_y")
          )
        );
      }
      reader.Close();
      return equip;
    }

    public string GetProjectId(string projectNo)
    {
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

    public void UpdateEquipment(Equipment equip)
    {
      string query =
        @"
          UPDATE electrical_equipment
          SET
          loc_x = @xLoc,
          loc_y = @yLoc,
          parent_distance = @parentDistance
          WHERE id = @equipId;
          ";
      OpenConnection();
      MySqlCommand command = new MySqlCommand(query, Connection);
      command.Parameters.AddWithValue("@xLoc", equip.loc.X);
      command.Parameters.AddWithValue("@yLoc", equip.loc.Y);
      command.Parameters.AddWithValue("@parentDistance", equip.parentDistance);
      command.Parameters.AddWithValue("@equipId", equip.id);
      command.ExecuteNonQuery();
    }

    public void UpdatePanel(Panel panel)
    {
      string query =
        @"
          UPDATE electrical_panels
          SET
          loc_x = @xLoc,
          loc_y = @yLoc,
          parent_distance = @parentDistance
          WHERE id = @equipId
          ";
      OpenConnection();
      MySqlCommand command = new MySqlCommand(query, Connection);
      command.Parameters.AddWithValue("@xLoc", panel.loc.X);
      command.Parameters.AddWithValue("@yLoc", panel.loc.Y);
      command.Parameters.AddWithValue("@parentDistance", panel.parentDistance);
      command.Parameters.AddWithValue("@equipId", panel.id);
      command.ExecuteNonQuery();
    }

    public void UpdateTransformer(Transformer xfmr)
    {
      string query =
        @"
          UPDATE electrical_transformers
          SET
          loc_x = @xLoc,
          loc_y = @yLoc,
          parent_distance = @parentDistance
          WHERE id = @equipId;
          ";
      OpenConnection();
      MySqlCommand command = new MySqlCommand(query, Connection);
      command.Parameters.AddWithValue("@xLoc", xfmr.loc.X);
      command.Parameters.AddWithValue("@yLoc", xfmr.loc.Y);
      command.Parameters.AddWithValue("@parentDistance", xfmr.parentDistance);
      command.Parameters.AddWithValue("@equipId", xfmr.id);
      command.ExecuteNonQuery();
    }
  }
}
