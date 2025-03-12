using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Accord.Statistics.Distributions;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using DocumentFormat.OpenXml.Office2010.CustomUI;
using ElectricalCommands.ElectricalEntity;
using ElectricalCommands.Equipment;
using ElectricalCommands.SingleLine;
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

    string GetSafeString(MySqlDataReader reader, string fieldName)
    {
      int index = reader.GetOrdinal(fieldName);
      if (!reader.IsDBNull(index))
      {
        return reader.GetString(index);
      }
      return string.Empty;
    }

    int GetSafeInt(MySqlDataReader reader, string fieldName)
    {
      int index = reader.GetOrdinal(fieldName);
      if (!reader.IsDBNull(index))
      {
        return reader.GetInt32(index);
      }
      return 0;
    }

    float GetSafeFloat(MySqlDataReader reader, string fieldName)
    {
      int index = reader.GetOrdinal(fieldName);
      if (!reader.IsDBNull(index))
      {
        return reader.GetFloat(index);
      }
      return 0;
    }

    bool GetSafeBoolean(MySqlDataReader reader, string fieldName)
    {
      int index = reader.GetOrdinal(fieldName);
      if (!reader.IsDBNull(index))
      {
        return reader.GetBoolean(index);
      }
      return false;
    }

    public List<Service> GetServices(string projectId)
    {
      List<Service> services = new List<Service>();
      string query =
        @"SELECT 
          electrical_services.id,
          electrical_service_amp_ratings.amp_rating,
          electrical_service_voltages.voltage,
          electrical_services.aic_rating,
          electrical_services.node_id,
          electrical_services.loc_x,
          electrical_services.loc_y,
          electrical_single_line_nodes.loc_x as node_x,
          electrical_single_line_nodes.loc_y as node_y,
          statuses.status
          FROM `electrical_services`
          LEFT JOIN electrical_service_meter_configs
          ON electrical_services.electrical_service_meter_config_id = electrical_service_meter_configs.id
          LEFT JOIN electrical_service_amp_ratings
          ON electrical_service_amp_ratings.id = electrical_services.electrical_service_amp_rating_id
          LEFT JOIN electrical_service_voltages
          ON electrical_service_voltages.id = electrical_services.electrical_service_voltage_id
          LEFT JOIN statuses
          ON statuses.id = electrical_services.status_id
          LEFT JOIN electrical_single_line_nodes 
          ON electrical_single_line_nodes.id = electrical_services.node_id
          WHERE electrical_services.project_id = @projectId";
      OpenConnection();
      MySqlCommand command = new MySqlCommand(query, Connection);
      command.Parameters.AddWithValue("projectId", projectId);
      MySqlDataReader reader = command.ExecuteReader();
      while (reader.Read())
      {
        services.Add(
          new Service(
            GetSafeString(reader, "id"),
            GetSafeString(reader, "node_id"),
            GetSafeString(reader, "status"),
            GetSafeInt(reader, "amp_rating"),
            GetSafeString(reader, "voltage"),
            GetSafeFloat(reader, "aic_rating"),
            new System.Drawing.Point(GetSafeInt(reader, "node_x"), GetSafeInt(reader, "node_y")),
            new Point3d(GetSafeFloat(reader, "loc_x"), GetSafeFloat(reader, "loc_y"), 0)
          )
        );
      }
      CloseConnection();
      reader.Close();
      return services;
    }

    public List<Meter> GetMeters(string projectId)
    {
      List<Meter> meters = new List<Meter>();
      string query =
        @"SELECT 
          electrical_meters.id,
          electrical_meters.has_cts,
          electrical_meters.is_space,
          electrical_meters.node_id,
          electrical_meters.aic_rating,
          electrical_single_line_nodes.loc_x,
          electrical_single_line_nodes.loc_y,
          statuses.status
          FROM `electrical_meters`
          LEFT JOIN statuses ON statuses.id = electrical_meters.status_id
          LEFT JOIN electrical_single_line_nodes 
          ON electrical_single_line_nodes.id = electrical_meters.node_id
          WHERE electrical_meters.project_id = @projectId";
      OpenConnection();
      MySqlCommand command = new MySqlCommand(query, Connection);
      command.Parameters.AddWithValue("projectId", projectId);
      MySqlDataReader reader = command.ExecuteReader();
      while (reader.Read())
      {
        meters.Add(
          new Meter(
            GetSafeString(reader, "id"),
            GetSafeString(reader, "node_id"),
            GetSafeString(reader, "status"),
            GetSafeBoolean(reader, "has_cts"),
            GetSafeBoolean(reader, "is_space"),
            GetSafeFloat(reader, "aic_rating"),
            new System.Drawing.Point(GetSafeInt(reader, "loc_x"), GetSafeInt(reader, "loc_y"))
          )
        );
      }
      CloseConnection();
      reader.Close();
      return meters;
    }

    public List<MainBreaker> GetMainBreakers(string projectId)
    {
      List<MainBreaker> mainBreakers = new List<MainBreaker>();
      string query =
        @"
        SELECT 
        electrical_main_breakers.id,
        electrical_main_breakers.node_id,
        electrical_main_breakers.has_ground_fault_protection,
        electrical_main_breakers.has_surge_protection,
        electrical_main_breakers.num_poles,
        electrical_main_breakers.aic_rating,
        electrical_single_line_nodes.loc_x,
        electrical_single_line_nodes.loc_y,
        statuses.status,
        electrical_service_amp_ratings.amp_rating
        FROM electrical_main_breakers
        LEFT JOIN statuses ON statuses.id = electrical_main_breakers.status_id
        LEFT JOIN electrical_service_amp_ratings ON electrical_service_amp_ratings.id = electrical_main_breakers.amp_rating_id
        LEFT JOIN electrical_single_line_nodes 
        ON electrical_single_line_nodes.id = electrical_main_breakers.node_id
        WHERE electrical_main_breakers.project_id = @projectId
        ";
      OpenConnection();
      MySqlCommand command = new MySqlCommand(query, Connection);
      command.Parameters.AddWithValue("projectId", projectId);
      MySqlDataReader reader = command.ExecuteReader();
      while (reader.Read())
      {
        mainBreakers.Add(
          new MainBreaker(
            GetSafeString(reader, "id"),
            GetSafeString(reader, "node_id"),
            GetSafeString(reader, "status"),
            GetSafeInt(reader, "amp_rating"),
            GetSafeBoolean(reader, "has_ground_fault_protection"),
            GetSafeBoolean(reader, "has_surge_protection"),
            GetSafeInt(reader, "num_poles"),
            GetSafeFloat(reader, "aic_rating"),
            new System.Drawing.Point(GetSafeInt(reader, "loc_x"), GetSafeInt(reader, "loc_y"))
          )
        );
      }
      CloseConnection();
      reader.Close();
      return mainBreakers;
    }

    public List<DistributionBus> GetDistributionBuses(string projectId)
    {
      List<DistributionBus> distributionBuses = new List<DistributionBus>();
      string query =
        @"SELECT
        electrical_distribution_buses.id,
        electrical_distribution_buses.node_id,
        electrical_distribution_buses.aic_rating,
        electrical_distribution_buses.loc_x,
        electrical_distribution_buses.loc_y,
        electrical_single_line_nodes.loc_x as node_x,
        electrical_single_line_nodes.loc_y as node_y,
        statuses.status,
        electrical_service_amp_ratings.amp_rating
        FROM electrical_distribution_buses
        LEFT JOIN electrical_service_amp_ratings 
        ON electrical_service_amp_ratings.id = electrical_distribution_buses.amp_rating_id
        LEFT JOIN statuses 
        ON statuses.id = electrical_distribution_buses.status_id
        LEFT JOIN electrical_single_line_nodes 
        ON electrical_single_line_nodes.id = electrical_distribution_buses.node_id
        WHERE electrical_distribution_buses.project_id = @projectId
        ";
      OpenConnection();
      MySqlCommand command = new MySqlCommand(query, Connection);
      command.Parameters.AddWithValue("projectId", projectId);
      MySqlDataReader reader = command.ExecuteReader();
      while (reader.Read())
      {
        distributionBuses.Add(
          new DistributionBus(
            GetSafeString(reader, "id"),
            GetSafeString(reader, "node_id"),
            GetSafeString(reader, "status"),
            GetSafeInt(reader, "amp_rating"),
            GetSafeFloat(reader, "aic_rating"),
            new System.Drawing.Point(GetSafeInt(reader, "node_x"), GetSafeInt(reader, "node_y")),
            new Point3d(GetSafeFloat(reader, "loc_x"), GetSafeFloat(reader, "loc_y"), 0)
          )
        );
      }
      CloseConnection();
      reader.Close();
      return distributionBuses;
    }

    public List<DistributionBreaker> GetDistributionBreakers(string projectId)
    {
      List<DistributionBreaker> distributionBreakers = new List<DistributionBreaker>();
      string query =
        @"SELECT
        electrical_distribution_breakers.id,
        electrical_distribution_breakers.node_id,
        electrical_distribution_breakers.aic_rating,
        electrical_distribution_breakers.num_poles,
        electrical_distribution_breakers.is_fuse_only,
        electrical_single_line_nodes.loc_x,
        electrical_single_line_nodes.loc_y,
        statuses.status,
        electrical_panel_bus_amp_ratings.amp_rating
        FROM electrical_distribution_breakers
        LEFT JOIN electrical_panel_bus_amp_ratings 
        ON electrical_panel_bus_amp_ratings.id = electrical_distribution_breakers.amp_rating_id
        LEFT JOIN statuses 
        ON statuses.id = electrical_distribution_breakers.status_id
        LEFT JOIN electrical_single_line_nodes 
        ON electrical_single_line_nodes.id = electrical_distribution_breakers.node_id
        WHERE electrical_distribution_breakers.project_id = @projectId
      ";
      OpenConnection();
      MySqlCommand command = new MySqlCommand(query, Connection);
      command.Parameters.AddWithValue("projectId", projectId);
      MySqlDataReader reader = command.ExecuteReader();
      while (reader.Read())
      {
        distributionBreakers.Add(
          new DistributionBreaker(
            GetSafeString(reader, "id"),
            GetSafeString(reader, "node_id"),
            GetSafeString(reader, "status"),
            GetSafeInt(reader, "amp_rating"),
            GetSafeInt(reader, "num_poles"),
            GetSafeBoolean(reader, "is_fuse_only"),
            GetSafeFloat(reader, "aic_rating"),
            new System.Drawing.Point(GetSafeInt(reader, "loc_x"), GetSafeInt(reader, "loc_y"))
          )
        );
      }
      CloseConnection();
      reader.Close();
      return distributionBreakers;
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
        electrical_panels.num_breakers,
        electrical_panels.circuit_no,
        electrical_panels.parent_distance,
        electrical_panels.loc_x,
        electrical_panels.loc_y,
        electrical_panels.is_distribution,
        electrical_panels.is_mlo,
        electrical_panels.is_recessed,
        electrical_panels.load_amperage,
        electrical_panels.kva,
        bus_amp_ratings.amp_rating as bus_amp_rating,
        main_amp_ratings.amp_rating as main_amp_rating,
        electrical_service_voltages.voltage,
        electrical_panels.aic_rating,
        electrical_panels.is_hidden_on_plan,
        electrical_panels.node_id,
        electrical_single_line_nodes.loc_x as node_x,
        electrical_single_line_nodes.loc_y as node_y,
        statuses.status
        FROM electrical_panels
        LEFT JOIN electrical_panel_bus_amp_ratings AS bus_amp_ratings
        ON bus_amp_ratings.id = electrical_panels.bus_amp_rating_id
        LEFT JOIN electrical_panel_bus_amp_ratings AS main_amp_ratings
        ON main_amp_ratings.id = electrical_panels.main_amp_rating_id
        LEFT JOIN electrical_service_voltages
        ON electrical_service_voltages.id = electrical_panels.voltage_id
        LEFT JOIN statuses
        ON statuses.id = electrical_panels.status_id
        LEFT JOIN electrical_single_line_nodes 
        ON electrical_single_line_nodes.id = electrical_panels.node_id
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
              reader.GetInt32("bus_amp_rating"),
              reader.GetInt32("main_amp_rating"),
              reader.GetBoolean("is_mlo"),
              reader.GetBoolean("is_recessed"),
              reader.GetString("voltage"),
              reader.GetFloat("load_amperage"),
              reader.GetFloat("kva"),
              reader.GetFloat("aic_rating"),
              reader.GetBoolean("is_hidden_on_plan"),
              reader.IsDBNull(reader.GetOrdinal("node_id")) ? string.Empty : reader.GetString("node_id"),
              reader.GetString("status"),
              new System.Drawing.Point(GetSafeInt(reader, "node_x"), GetSafeInt(reader, "node_y")),
              reader.GetInt32("num_breakers"),
              reader.GetInt32("circuit_no")
            )
          );
      }
      CloseConnection();
      reader.Close();
      return panels;
    }

    public List<PanelBreaker> GetPanelBreakers(string projectId)
    {
      List<PanelBreaker> panelBreakers = new List<PanelBreaker>();
      string query =
        @"
        SELECT 
        electrical_panel_breakers.id,
        electrical_panel_breakers.node_id,
        electrical_panel_breakers.num_poles,
        electrical_panel_breakers.circuit_no,
        electrical_panel_breakers.aic_rating,
        electrical_panel_bus_amp_ratings.amp_rating,
        electrical_single_line_nodes.loc_x,
        electrical_single_line_nodes.loc_y,
        statuses.status
        FROM electrical_panel_breakers
        LEFT JOIN statuses ON statuses.id = electrical_panel_breakers.status_id
        LEFT JOIN electrical_single_line_nodes ON electrical_single_line_nodes.id = electrical_panel_breakers.node_id
        LEFT JOIN electrical_panel_bus_amp_ratings ON electrical_panel_bus_amp_ratings.id = electrical_panel_breakers.amp_rating_id
        WHERE electrical_panel_breakers.project_id = @projectId
        ";
      OpenConnection();
      MySqlCommand command = new MySqlCommand(query, Connection);
      command.Parameters.AddWithValue("projectId", projectId);
      MySqlDataReader reader = command.ExecuteReader();
      while (reader.Read())
      {
        panelBreakers.Add(
          new PanelBreaker(
            GetSafeString(reader, "id"),
            GetSafeString(reader, "node_id"),
            GetSafeString(reader, "status"),
            GetSafeInt(reader, "amp_rating"),
            GetSafeInt(reader, "num_poles"),
            GetSafeInt(reader, "circuit_no"),
            GetSafeFloat(reader, "aic_rating"),
            new System.Drawing.Point(GetSafeInt(reader, "loc_x"), GetSafeInt(reader, "loc_y"))
          )
        );
      }
      CloseConnection();
      reader.Close();
      return panelBreakers;
    }

    public List<Disconnect> GetDisconnects(string projectId)
    {
      List<Disconnect> disconnects = new List<Disconnect>();
      string query =
        @"
        SELECT 
        electrical_disconnects.id,
        electrical_disconnects.node_id,
        electrical_disconnects.parent_id,
        electrical_disconnect_as_sizes.as_size,
        electrical_disconnect_af_sizes.af_size,
        electrical_disconnects.num_poles,
        electrical_disconnects.parent_distance,
        electrical_disconnects.aic_rating,
        electrical_disconnects.loc_x,
        electrical_disconnects.loc_y,
        electrical_single_line_nodes.loc_x as node_x,
        electrical_single_line_nodes.loc_y as node_y,
        statuses.status
        FROM electrical_disconnects
        LEFT JOIN statuses ON statuses.id = electrical_disconnects.status_id
        LEFT JOIN electrical_disconnect_as_sizes ON electrical_disconnect_as_sizes.id = electrical_disconnects.as_size_id
        LEFT JOIN electrical_disconnect_af_sizes ON electrical_disconnect_af_sizes.id = electrical_disconnects.af_size_id
        LEFT JOIN electrical_single_line_nodes ON electrical_single_line_nodes.id = electrical_disconnects.node_id        
        WHERE electrical_disconnects.project_id = @projectId
        ";
      OpenConnection();
      MySqlCommand command = new MySqlCommand(query, Connection);
      command.Parameters.AddWithValue("projectId", projectId);
      MySqlDataReader reader = command.ExecuteReader();
      while (reader.Read())
      {
        disconnects.Add(
          new Disconnect(
            GetSafeString(reader, "id"),
            GetSafeString(reader, "parent_id"),
            GetSafeInt(reader, "parent_distance"),
            GetSafeString(reader, "node_id"),
            GetSafeString(reader, "status"),
            GetSafeInt(reader, "as_size"),
            GetSafeInt(reader, "af_size"),
            GetSafeInt(reader, "num_poles"),
            GetSafeFloat(reader, "aic_rating"),
            new System.Drawing.Point(GetSafeInt(reader, "node_x"), GetSafeInt(reader, "node_y")),
            new Point3d(GetSafeFloat(reader, "loc_x"), GetSafeFloat(reader, "loc_y"), 0)
          )
        );
      }
      CloseConnection();
      reader.Close();
      return disconnects;
    }

    public List<Transformer> GetTransformers(string projectId)
    {
      List<Transformer> xfmrs = new List<Transformer>();
      string query =
        @"
        SELECT 
        electrical_transformers.id,
        electrical_transformers.parent_id,
        electrical_transformers.name,
        electrical_transformers.circuit_no,
        electrical_transformers.parent_distance,
        electrical_transformers.loc_x,
        electrical_transformers.loc_y,
        electrical_transformer_kva_ratings.kva_rating,
        electrical_transformer_voltages.voltage,
        electrical_transformers.aic_rating,
        electrical_transformers.is_hidden_on_plan,
        electrical_transformers.node_id,
        electrical_single_line_nodes.loc_x as node_x,
        electrical_single_line_nodes.loc_y as node_y,
        statuses.status
        FROM electrical_transformers
        LEFT JOIN electrical_transformer_kva_ratings
        ON electrical_transformer_kva_ratings.id = electrical_transformers.kva_id
        LEFT JOIN electrical_transformer_voltages
        ON electrical_transformer_voltages.id = electrical_transformers.voltage_id
        LEFT JOIN statuses
        ON statuses.id = electrical_transformers.status_id
        LEFT JOIN electrical_single_line_nodes
        ON electrical_single_line_nodes.id = electrical_transformers.node_id
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
            GetSafeString(reader, "id"),
            GetSafeString(reader, "parent_id"),
            GetSafeString(reader, "name"),
            GetSafeInt(reader, "parent_distance"),
            GetSafeFloat(reader, "loc_x"),
            GetSafeFloat(reader, "loc_y"),
            GetSafeFloat(reader, "kva_rating"),
            GetSafeString(reader, "voltage"),
            GetSafeFloat(reader, "aic_rating"),
            GetSafeBoolean(reader, "is_hidden_on_plan"),
            GetSafeString(reader, "node_id"),
            GetSafeString(reader, "status"),
            new System.Drawing.Point(GetSafeInt(reader, "node_x"), GetSafeInt(reader, "node_y")),
            reader.GetInt32("circuit_no")
          )
        );
      }
      CloseConnection();
      reader.Close();
      return xfmrs;
    }

    public List<NodeLink> GetNodeLinks(string projectId)
    {
      List<NodeLink> nodeLinks = new List<NodeLink>();
      string query =
        @"
        SELECT * FROM electrical_single_line_node_links WHERE project_id = @projectId
        ";
      OpenConnection();
      MySqlCommand command = new MySqlCommand(query, Connection);
      command.Parameters.AddWithValue("projectId", projectId);
      MySqlDataReader reader = command.ExecuteReader();
      while (reader.Read())
      {
        nodeLinks.Add(
          new NodeLink(
            GetSafeString(reader, "id"),
            GetSafeString(reader, "input_connector_node_id"),
            GetSafeString(reader, "output_connector_node_id")
          )
        );
      }
      CloseConnection();
      reader.Close();
      return nodeLinks;
    }

    public List<GroupNode> GetGroupNodes(string projectId)
    {
      List<GroupNode> groupNodes = new List<GroupNode>();
      string query =
        @"
        SELECT 
        electrical_single_line_groups.id,
        electrical_single_line_groups.width,
        electrical_single_line_groups.height,
        electrical_single_line_groups.name,
        electrical_single_line_groups.loc_x,
        electrical_single_line_groups.loc_y,
        statuses.status
        FROM electrical_single_line_groups
        LEFT JOIN statuses ON statuses.id = electrical_single_line_groups.status_id
        WHERE project_id = @projectId
        ORDER BY electrical_single_line_groups.loc_x ASC
        ";
      OpenConnection();
      MySqlCommand command = new MySqlCommand(query, Connection);
      command.Parameters.AddWithValue("projectId", projectId);
      MySqlDataReader reader = command.ExecuteReader();
      while (reader.Read())
      {
        groupNodes.Add(
          new GroupNode(
            GetSafeString(reader, "id"),
            GetSafeInt(reader, "width"),
            GetSafeInt(reader, "height"),
            GetSafeString(reader, "name"),
            new System.Drawing.Point(GetSafeInt(reader, "loc_x"), GetSafeInt(reader, "loc_y")),
            GetSafeString(reader, "status")
          )
        );
      }
      CloseConnection();
      reader.Close();
      return groupNodes;
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
        electrical_equipment.mca,
        electrical_equipment.hp,
        electrical_equipment.mounting_height,
        electrical_equipment.circuit_no,
        electrical_equipment.has_plug,
        electrical_equipment.is_hidden_on_plan
        FROM electrical_equipment
        LEFT JOIN electrical_panels
        ON electrical_panels.id = electrical_equipment.parent_id
        LEFT JOIN electrical_equipment_categories
        ON electrical_equipment.category_id = electrical_equipment_categories.id
        LEFT JOIN electrical_equipment_voltages
        ON electrical_equipment_voltages.id = electrical_equipment.voltage_id
        WHERE electrical_equipment.project_id = @projectId
        ORDER BY electrical_equipment.equip_no ASC";
      this.OpenConnection();
      MySqlCommand command = new MySqlCommand(query, Connection);
      command.Parameters.AddWithValue("projectId", projectId);
      MySqlDataReader reader = command.ExecuteReader();
      while (reader.Read())
      {
        bool is3Phase = reader.GetInt32("is_three_phase") == 1;
        try
        {
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
              reader.GetFloat("mca"),
              reader.GetString("hp"),
              reader.GetInt32("mounting_height"),
              reader.GetInt32("circuit_no"),
              reader.GetBoolean("has_plug"),
              reader.GetBoolean("is_hidden_on_plan")
            )
          );
        }
        catch { }
      }
      CloseConnection();
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
        electrical_lighting.circuit_no,
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
        symbols.label_transform_h_x,
        symbols.label_transform_h_y,
        symbols.label_transform_v_x,
        symbols.label_transform_v_y,
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
            reader.IsDBNull(reader.GetOrdinal("name")) ? string.Empty : reader.GetString("name"),
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
            reader.GetBoolean("rotate"),
            reader.GetFloat("paper_space_scale"),
            reader.GetBoolean("em_capable"),
            reader.GetFloat("label_transform_h_x"),
            reader.GetFloat("label_transform_h_y"),
            reader.GetFloat("label_transform_v_x"),
            reader.GetFloat("label_transform_v_y"),
            reader.GetInt32("circuit_no")
          )
        );
      }
      CloseConnection();
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
            reader.GetBoolean("occupancy")
          )
        );
      }
      CloseConnection();
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
    public void UpdateEquipment(Equipment equip) {
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
      command.Parameters.AddWithValue("@xLoc", equip.Location.X);
      command.Parameters.AddWithValue("@yLoc", equip.Location.Y);
      command.Parameters.AddWithValue("@parentDistance", equip.ParentDistance);
      command.Parameters.AddWithValue("@equipId", equip.Id);
      command.ExecuteNonQuery();
      CloseConnection();
    }

    public void InsertLightingEquipment(List<string> lightings, string panelId, int circuitNo, string projectId) {

      float newWattage = 0;
      string query =
        @"
        SELECT
        electrical_lighting.wattage
        FROM electrical_lighting
        WHERE electrical_lighting.id = @id";

      this.OpenConnection();
      foreach (var id in lightings) {
        MySqlCommand command = new MySqlCommand(query, Connection);
        command.Parameters.AddWithValue("@id", id);
        MySqlDataReader reader = command.ExecuteReader();
        while (reader.Read()) {
          newWattage += reader.GetInt32("wattage");
        }
        reader.Close();
      }

      query =
        @"INSERT INTO electrical_equipment (id, project_id, parent_id, description, category_id, voltage_id, 
        fla, is_three_phase, circuit_no, spec_sheet_from_client, aic_rating, color_code, connection_type_id, va, load_type) VALUES (@id, @projectId, @parentId, @description, @category, 
        @voltage, @fla, @isThreePhase, @circuit, @specFromClient, @aicRating, @colorCode, @connectionId, @va, @loadType)";

      MySqlCommand  command2 = new MySqlCommand(query, Connection);
      command2.Parameters.AddWithValue("@id", Guid.NewGuid().ToString());
      command2.Parameters.AddWithValue("@projectId", projectId);
      command2.Parameters.AddWithValue("@parentId", panelId);
      command2.Parameters.AddWithValue("@description", "Lighting");
      command2.Parameters.AddWithValue("@category", 5);
      command2.Parameters.AddWithValue("@voltage", 1);
      command2.Parameters.AddWithValue("@fla", Math.Round(newWattage / 115, 1, MidpointRounding.AwayFromZero));
      command2.Parameters.AddWithValue("@va", newWattage);
      command2.Parameters.AddWithValue("@isThreePhase", false);
      command2.Parameters.AddWithValue("@circuit", circuitNo);
      command2.Parameters.AddWithValue("@specFromClient", false);
      command2.Parameters.AddWithValue("@aicRating", 0);
      command2.Parameters.AddWithValue("@colorCode", "#FF00FF");
      command2.Parameters.AddWithValue("@connectionId", 1);
      command2.Parameters.AddWithValue("@loadType", 3);
      command2.ExecuteNonQuery();
      CloseConnection();
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
      command.Parameters.AddWithValue("@xLoc", panel.Location.X);
      command.Parameters.AddWithValue("@yLoc", panel.Location.Y);
      command.Parameters.AddWithValue("@parentDistance", panel.ParentDistance);
      command.Parameters.AddWithValue("@equipId", panel.Id);
      command.ExecuteNonQuery();
      CloseConnection();
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
      CloseConnection();
    }

    public void UpdateTransformer(Transformer xfmr)
    {
      string query =
        @"
          UPDATE electrical_transformers
          SET
          loc_x = @xLoc,
          loc_y = @yLoc,
          aic_rating = @aicRating,
          parent_distance = @parentDistance
          WHERE id = @equipId;
          ";
      OpenConnection();
      MySqlCommand command = new MySqlCommand(query, Connection);
      command.Parameters.AddWithValue("@xLoc", xfmr.Location.X);
      command.Parameters.AddWithValue("@yLoc", xfmr.Location.Y);
      command.Parameters.AddWithValue("@aicRating", xfmr.AicRating);
      command.Parameters.AddWithValue("@parentDistance", xfmr.ParentDistance);
      command.Parameters.AddWithValue("@equipId", xfmr.Id);
      command.ExecuteNonQuery();
      CloseConnection();
    }

    public void UpdatePlaceable(PlaceableElectricalEntity placeable)
    {
      string query = $"UPDATE {placeable.TableName}";
      if (placeable.NodeType != NodeType.Service)
      {
        query +=
          @"
          SET
          loc_x = @xLoc,
          loc_y = @yLoc,
          aic_rating = @aicRating,
          parent_distance = @parentDistance
          WHERE id = @placeableId
          ";
      }
      else
      {
        query +=
          @"
          SET
          loc_x = @xLoc,
          loc_y = @yLoc
          WHERE id = @placeableId
          ";
      }
      OpenConnection();
      MySqlCommand command = new MySqlCommand(query, Connection);
      command.Parameters.AddWithValue("@xLoc", placeable.Location.X);
      command.Parameters.AddWithValue("@yLoc", placeable.Location.Y);
      if (placeable.NodeType != NodeType.Service)
      {
        command.Parameters.AddWithValue("@aicRating", placeable.AicRating);
        command.Parameters.AddWithValue("@parentDistance", placeable.ParentDistance);
      }
      command.Parameters.AddWithValue("@placeableId", placeable.Id);
      command.ExecuteNonQuery();
      CloseConnection();
    }
  }
}
