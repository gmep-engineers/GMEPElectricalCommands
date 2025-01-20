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
          electrical_service_voltages.voltage,
          electrical_services.aic_rating
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
            reader.GetString("voltage"),
            reader.GetFloat("aic_rating")
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
        electrical_service_voltages.voltage,
        electrical_panels.aic_rating
        FROM electrical_panels
        LEFT JOIN electrical_panel_bus_amp_ratings
        ON electrical_panel_bus_amp_ratings.id = electrical_panels.bus_amp_rating_id
        LEFT JOIN electrical_service_voltages
        ON electrical_service_voltages.id = electrical_panels.voltage_id
        WHERE electrical_panels.project_id = @projectId
        ORDER BY electrical_panels.name ASC";
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
            reader.GetString("voltage"),
            reader.GetFloat("aic_rating")
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
        WHERE electrical_transformers.project_id = @projectId
        ORDER BY electrical_transformers.name ASC";
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
            reader.GetString("voltage"),
            reader.GetFloat("aic_rating")
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
        electrical_equipment.fla,
        electrical_equipment.is_three_phase,
        electrical_equipment.parent_distance,
        electrical_equipment.loc_x,
        electrical_equipment.loc_y,
        electrical_equipment.mca_id,
        electrical_equipment_mca_ratings.mca_rating,
        electrical_equipment.hp,
        electrical_equipment.mounting_height
        FROM electrical_equipment
        LEFT JOIN electrical_panels
        ON electrical_panels.id = electrical_equipment.parent_id
        LEFT JOIN electrical_equipment_categories
        ON electrical_equipment.category_id = electrical_equipment_categories.id
        LEFT JOIN electrical_equipment_voltages
        ON electrical_equipment_voltages.id = electrical_equipment.voltage_id
        LEFT JOIN electrical_equipment_mca_ratings
        ON electrical_equipment_mca_ratings.id = electrical_equipment.mca_id
        WHERE electrical_equipment.project_id = @projectId
        ORDER BY electrical_equipment.equip_no ASC";
      this.OpenConnection();
      MySqlCommand command = new MySqlCommand(query, Connection);
      command.Parameters.AddWithValue("projectId", projectId);
      MySqlDataReader reader = command.ExecuteReader();
      while (reader.Read())
      {
        bool is3Phase = reader.GetInt32("is_three_phase") == 1;
        int mcaId = reader.GetInt32("mca_id");
        int mca = -1;
        if (mcaId != 0)
        {
          mca = reader.GetInt32("mca_rating");
        }
        equip.Add(
          new Equipment(
            reader.GetString("id"),
            reader.GetString("parent_id"),
            reader.GetString("name"),
            reader.GetString("equip_no"),
            reader.GetString("description"),
            reader.GetString("category"),
            reader.GetInt32("voltage"),
            reader.GetFloat("fla"),
            is3Phase,
            reader.GetInt32("parent_distance"),
            reader.GetFloat("loc_x"),
            reader.GetFloat("loc_y"),
            mca,
            reader.GetString("hp"),
            reader.GetInt32("mounting_height")
          )
        );
      }
      reader.Close();
      return equip;
    }

    public List<LightingFixture> GetLightingFixtures(string projectId)
    {
      List<LightingFixture> ltg = new List<LightingFixture>();
      string query =
        @"
        SELECT
        electrical_lighting.id,
        electrical_lighting.parent_id,
        electrical_panels.name,
        electrical_lighting.control_id,
        electrical_lighting.description,
        electrical_equipment_voltages.voltage,
        electrical_lighting.wattage,
        electrical_lighting.em_capable,
        electrical_lighting.model_no,
        electrical_lighting.tag,
        electrical_lighting.qty,
        electrical_lighting.manufacturer,
        symbols.block_name,
        symbols.rotate,
        symbols.paper_space_scale,
        electrical_lighting.notes,
        electrical_lighting_mounting_types.mounting
        FROM electrical_lighting
        LEFT JOIN electrical_panels
        ON electrical_panels.id = electrical_lighting.parent_id
        LEFT JOIN electrical_equipment_voltages
        ON electrical_equipment_voltages.id = electrical_lighting.voltage_id
        LEFT JOIN symbols
        ON symbols.id = electrical_lighting.symbol_id
        LEFT JOIN electrical_lighting_mounting_types
        ON electrical_lighting_mounting_types.id = electrical_lighting.mounting_type_id
        WHERE electrical_lighting.project_id = @projectId
        ORDER BY electrical_lighting.tag ASC";
      this.OpenConnection();
      MySqlCommand command = new MySqlCommand(query, Connection);
      command.Parameters.AddWithValue("projectId", projectId);
      MySqlDataReader reader = command.ExecuteReader();
      while (reader.Read())
      {
        ltg.Add(
          new LightingFixture(
            reader.GetString("id"),
            reader.GetString("parent_id"),
            reader.GetString("name"),
            reader.GetString("tag"),
            reader.GetString("control_id"),
            reader.GetString("block_name"),
            reader.GetInt32("voltage"),
            reader.GetFloat("wattage"),
            reader.GetString("description"),
            reader.GetInt32("qty"),
            reader.GetString("mounting"),
            reader.GetString("manufacturer"),
            reader.GetString("model_no"),
            reader.GetString("notes"),
            reader.GetInt32("rotate"),
            reader.GetFloat("paper_space_scale"),
            reader.GetInt32("em_capable")
          )
        );
      }
      reader.Close();
      return ltg;
    }

    public List<LightingControl> GetLightingControls(string projectId)
    {
      List<LightingControl> ltgCtrl = new List<LightingControl>();
      string query =
        @"
        SELECT
        electrical_lighting_controls.id,
        electrical_lighting_controls.name,
        electrical_lighting_control_types.control,
        electrical_lighting_controls.occupancy
        FROM electrical_lighting_controls
        LEFT JOIN electrical_lighting_control_types
        ON electrical_lighting_control_types.id = electrical_lighting_controls.control_type_id
        WHERE project_id = @projectId
        ORDER BY name ASC";
      this.OpenConnection();
      MySqlCommand command = new MySqlCommand(query, Connection);
      command.Parameters.AddWithValue("projectId", projectId);
      MySqlDataReader reader = command.ExecuteReader();
      while (reader.Read())
      {
        ltgCtrl.Add(
          new LightingControl(
            reader.GetString("id"),
            reader.GetString("name"),
            reader.GetString("control_type"),
            reader.GetInt32("occupancy")
          )
        );
      }
      reader.Close();
      return ltgCtrl;
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

    public void UpdatePanelAic(string id, double aicRating)
    {
      string query =
        @"
          UPDATE electrical_panels
          SET
          aic_rating = @aicRating
          WHERE id = @equipId
          ";
      OpenConnection();
      MySqlCommand command = new MySqlCommand(query, Connection);
      command.Parameters.AddWithValue("@aicRating", aicRating);
      command.Parameters.AddWithValue("@equipId", id);
      command.ExecuteNonQuery();
    }

    public void UpdateTransformerAic(string id, double aicRating)
    {
      string query =
        @"
          UPDATE electrical_transformers
          SET
          aic_rating = @aicRating
          WHERE id = @equipId
          ";
      OpenConnection();
      MySqlCommand command = new MySqlCommand(query, Connection);
      command.Parameters.AddWithValue("@aicRating", aicRating);
      command.Parameters.AddWithValue("@equipId", id);
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
