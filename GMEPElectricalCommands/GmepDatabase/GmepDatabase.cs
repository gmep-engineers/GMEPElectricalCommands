using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Accord.Statistics.Distributions;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using DocumentFormat.OpenXml.Office2010.CustomUI;
using DocumentFormat.OpenXml.Office2010.Excel;
using ElectricalCommands;
using ElectricalCommands.ElectricalEntity;
using ElectricalCommands.Equipment;
using ElectricalCommands.Notes;
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

    public MySqlConnection GetConnection()
    {
      return Connection;
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
        electrical_distribution_buses.parent_id,
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
            GetSafeString(reader, "parent_id"),
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

    public List<PanelNote> GetPanelNotes(string panelId)
    {
      List<PanelNote> panelNotes = new List<PanelNote>();
      string query =
        @"
        SELECT
        electrical_panel_note_panel_rel.panel_id,
        electrical_panel_note_panel_rel.circuit_no,
        electrical_panel_note_panel_rel.length,
        electrical_panel_note_panel_rel.note_id,
        electrical_panel_notes.note
        FROM electrical_panel_note_panel_rel
        LEFT JOIN electrical_panel_notes ON electrical_panel_notes.id = electrical_panel_note_panel_rel.note_id
        WHERE electrical_panel_note_panel_rel.panel_id = @panelId
        ORDER BY electrical_panel_note_panel_rel.note_id, electrical_panel_notes.date
      ";
      OpenConnection();
      MySqlCommand command = new MySqlCommand(query, Connection);
      command.Parameters.AddWithValue("panelId", panelId);
      MySqlDataReader reader = command.ExecuteReader();
      List<string> noteIds = new List<string>();
      int number = 0;
      while (reader.Read())
      {
        string noteId = GetSafeString(reader, "note_id");
        if (!noteIds.Contains(noteId))
        {
          number++;
          noteIds.Add(noteId);
        }
        panelNotes.Add(
          new PanelNote(
            number,
            GetSafeString(reader, "panel_id"),
            GetSafeInt(reader, "circuit_no"),
            GetSafeInt(reader, "length"),
            GetSafeString(reader, "note").ToUpper()
          )
        );
      }
      CloseConnection();
      reader.Close();
      return panelNotes;
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
        electrical_panels.location,
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
            GetSafeString(reader, "id"),
            GetSafeString(reader, "parent_id"),
            GetSafeString(reader, "name"),
            GetSafeInt(reader, "parent_distance"),
            GetSafeFloat(reader, "loc_x"),
            GetSafeFloat(reader, "loc_y"),
            GetSafeInt(reader, "bus_amp_rating"),
            GetSafeInt(reader, "main_amp_rating"),
            GetSafeBoolean(reader, "is_mlo"),
            GetSafeBoolean(reader, "is_recessed"),
            GetSafeString(reader, "voltage"),
            GetSafeFloat(reader, "load_amperage"),
            GetSafeFloat(reader, "kva"),
            GetSafeFloat(reader, "aic_rating"),
            GetSafeBoolean(reader, "is_hidden_on_plan"),
            GetSafeString(reader, "node_id"),
            GetSafeString(reader, "status"),
            new System.Drawing.Point(GetSafeInt(reader, "node_x"), GetSafeInt(reader, "node_y")),
            GetSafeInt(reader, "num_breakers"),
            GetSafeInt(reader, "circuit_no"),
            GetSafeString(reader, "location")
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

    public int GetNumDuplex(string equipId)
    {
      int numDuplex = 0;
      string query =
        @"
      SELECT num_conv_duplex FROM electrical_equipment WHERE id = @equipId
      ";
      OpenConnection();
      MySqlCommand command = new MySqlCommand(query, Connection);
      command.Parameters.AddWithValue("equipId", equipId);
      MySqlDataReader reader = command.ExecuteReader();
      if (reader.Read())
      {
        numDuplex = GetSafeInt(reader, "num_conv_duplex");
      }
      CloseConnection();
      reader.Close();
      return numDuplex;
    }

    public (double, int, bool) GetEquipmentPowerSpecs(string equipId)
    {
      double fla = 0;
      int voltage = 0;
      bool is3Phase = false;
      string query =
        @"
        SELECT fla, voltage, is_three_phase
        FROM electrical_equipment
        LEFT JOIN electrical_equipment_voltages 
        ON electrical_equipment_voltages.id = electrical_equipment.voltage_id
        WHERE electrical_equipment.id = @equipId
      ";
      OpenConnection();
      MySqlCommand command = new MySqlCommand(query, Connection);
      command.Parameters.AddWithValue("equipId", equipId);
      MySqlDataReader reader = command.ExecuteReader();
      if (reader.Read())
      {
        fla = GetSafeFloat(reader, "fla");
        voltage = GetSafeInt(reader, "voltage");
        is3Phase = GetSafeBoolean(reader, "is_three_phase");
      }
      CloseConnection();
      reader.Close();
      return (fla, voltage, is3Phase);
    }

    public List<Equipment> GetEquipment(string projectId, bool singleLineOnly = false)
    {
      List<Equipment> equip = new List<Equipment>();
      string query =
        @"
        SELECT
        electrical_equipment.id,
        electrical_equipment.parent_id,
        electrical_panels.name,
        electrical_equipment_connection_symbols.connection_name,
        electrical_equipment.equip_no,
        electrical_equipment_categories.category,
        electrical_equipment.description,
        electrical_equipment_voltages.voltage,
        electrical_equipment.fla,
        electrical_equipment.is_three_phase,
        electrical_equipment.parent_distance,
        electrical_equipment.loc_x,
        electrical_equipment.loc_y,
        electrical_disconnect_af_sizes.af_size as mocp,
        electrical_equipment.hp,
        electrical_equipment.va,
        electrical_equipment.mounting_height,
        electrical_equipment.circuit_no,
        electrical_equipment.has_plug,
        electrical_equipment.node_id,
        electrical_equipment.is_hidden_on_plan,
        electrical_equipment.circuit_half,
        electrical_equipment.phase_a_va,
        electrical_equipment.phase_b_va,
        electrical_equipment.phase_c_va,
        statuses.status
        FROM electrical_equipment
        LEFT JOIN electrical_panels
        ON electrical_panels.id = electrical_equipment.parent_id
        LEFT JOIN electrical_equipment_categories
        ON electrical_equipment.category_id = electrical_equipment_categories.id
        LEFT JOIN electrical_equipment_voltages
        ON electrical_equipment_voltages.id = electrical_equipment.voltage_id
        LEFT JOIN statuses ON statuses.id = electrical_equipment.status_id
        LEFT JOIN electrical_equipment_connection_symbols ON electrical_equipment_connection_symbols.id = electrical_equipment.connection_symbol_id
        LEFT JOIN electrical_disconnect_af_sizes on electrical_disconnect_af_sizes.id = electrical_equipment.mocp_id
        WHERE electrical_equipment.project_id = @projectId";
      if (singleLineOnly)
      {
        query +=
          @"
          AND electrical_equipment.node_id IS NOT NULL        
          ";
      }
      query +=
        @"
        ORDER BY electrical_equipment.equip_no ASC    
        ";
      this.OpenConnection();
      MySqlCommand command = new MySqlCommand(query, Connection);
      command.Parameters.AddWithValue("projectId", projectId);
      MySqlDataReader reader = command.ExecuteReader();
      while (reader.Read())
      {
        equip.Add(
          new Equipment(
            GetSafeString(reader, "id"),
            GetSafeString(reader, "node_id"),
            GetSafeString(reader, "parent_id"),
            GetSafeString(reader, "name"),
            GetSafeString(reader, "equip_no"),
            GetSafeString(reader, "description"),
            GetSafeString(reader, "category"),
            GetSafeInt(reader, "voltage"),
            GetSafeFloat(reader, "fla"),
            GetSafeBoolean(reader, "is_three_phase"),
            GetSafeInt(reader, "parent_distance"),
            GetSafeFloat(reader, "loc_x"),
            GetSafeFloat(reader, "loc_y"),
            GetSafeInt(reader, "mocp"),
            GetSafeString(reader, "hp"),
            GetSafeInt(reader, "va"),
            GetSafeInt(reader, "mounting_height"),
            GetSafeInt(reader, "circuit_no"),
            GetSafeBoolean(reader, "has_plug"),
            GetSafeBoolean(reader, "is_hidden_on_plan"),
            GetSafeString(reader, "status"),
            GetSafeString(reader, "connection_name"),
            GetSafeInt(reader, "circuit_half"),
            GetSafeFloat(reader, "phase_a_va"),
            GetSafeFloat(reader, "phase_b_va"),
            GetSafeFloat(reader, "phase_c_va")
          )
        );
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
        electrical_lighting.location_id,
        electrical_lighting.circuit_no,
        electrical_equipment_voltages.voltage,
        electrical_lighting.wattage,
        electrical_lighting.em_capable,
        electrical_lighting.model_no,
        electrical_lighting.tag,
        electrical_lighting.qty,
        electrical_lighting.manufacturer,
        electrical_lighting.occupancy,
        electrical_lighting.has_photocell,
        symbols.block_name,
        symbols.rotate,
        symbols.paper_space_scale,
        symbols.label_transform_h_x,
        symbols.label_transform_h_y,
        symbols.label_transform_v_x,
        symbols.label_transform_v_y,
        electrical_lighting.notes,
        electrical_lighting_mounting_types.mounting,
        electrical_lighting_driver_types.driver_type
        FROM electrical_lighting
        LEFT JOIN electrical_panels
        ON electrical_panels.id = electrical_lighting.parent_id
        LEFT JOIN electrical_equipment_voltages
        ON electrical_equipment_voltages.id = electrical_lighting.voltage_id
        LEFT JOIN electrical_lighting_symbols as symbols
        ON symbols.id = electrical_lighting.symbol_id
        LEFT JOIN electrical_lighting_mounting_types
        ON electrical_lighting_mounting_types.id = electrical_lighting.mounting_type_id
        LEFT JOIN electrical_lighting_driver_types
        ON electrical_lighting_driver_types.id = electrical_lighting.driver_type_id
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
            GetSafeString(reader, "id"),
            GetSafeString(reader, "parent_id"),
            GetSafeString(reader, "location_id"),
            GetSafeString(reader, "name"),
            GetSafeString(reader, "tag"),
            GetSafeString(reader, "control_id"),
            GetSafeString(reader, "block_name"),
            GetSafeInt(reader, "voltage"),
            Math.Round(GetSafeFloat(reader, "wattage"), 1),
            GetSafeString(reader, "description"),
            GetSafeInt(reader, "qty"),
            GetSafeString(reader, "mounting"),
            GetSafeString(reader, "manufacturer"),
            GetSafeString(reader, "model_no"),
            GetSafeString(reader, "notes"),
            GetSafeBoolean(reader, "rotate"),
            GetSafeFloat(reader, "paper_space_scale"),
            GetSafeBoolean(reader, "em_capable"),
            GetSafeBoolean(reader, "has_photocell"),
            GetSafeBoolean(reader, "occupancy"),
            GetSafeFloat(reader, "label_transform_h_x"),
            GetSafeFloat(reader, "label_transform_h_y"),
            GetSafeFloat(reader, "label_transform_v_x"),
            GetSafeFloat(reader, "label_transform_v_y"),
            GetSafeInt(reader, "circuit_no"),
            GetSafeString(reader, "driver_type")
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
        electrical_lighting_driver_types.driver_type,
        electrical_lighting_controls.occupancy
        FROM electrical_lighting_controls
        LEFT JOIN electrical_lighting_driver_types
        ON electrical_lighting_driver_types.id = electrical_lighting_controls.driver_type_id
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
            GetSafeString(reader, "id"),
            GetSafeString(reader, "name"),
            GetSafeString(reader, "driver_type"),
            GetSafeBoolean(reader, "occupancy")
          )
        );
      }
      CloseConnection();
      reader.Close();
      return ltgCtrl;
    }

    public List<LightingLocation> GetLightingLocations(string projectId)
    {
      List<LightingLocation> locations = new List<LightingLocation>();
      string query =
        @"
        SELECT * FROM electrical_lighting_locations WHERE project_id = @projectId
        ";
      OpenConnection();
      MySqlCommand command = new MySqlCommand(query, Connection);
      command.Parameters.AddWithValue("projectId", projectId);
      MySqlDataReader reader = command.ExecuteReader();
      while (reader.Read())
      {
        locations.Add(
          new LightingLocation(
            GetSafeString(reader, "id"),
            GetSafeString(reader, "location"),
            GetSafeBoolean(reader, "outdoor"),
            GetSafeString(reader, "timeclock_id")
          )
        );
      }
      CloseConnection();
      reader.Close();
      return locations;
    }

    public List<LightingTimeClock> GetLightingTimeClocks(string projectId)
    {
      List<LightingTimeClock> clocks = new List<LightingTimeClock>();
      string query =
        @"
        SELECT 
        electrical_lighting_timeclocks.id as timeclock_id,
        electrical_lighting_timeclocks.name,
        electrical_lighting_timeclocks.bypass_switch_name,
        electrical_lighting_timeclocks.bypass_switch_location,
        electrical_lighting_timeclocks.adjacent_panel_id,
        electrical_equipment_voltages.voltage
        FROM electrical_lighting_timeclocks 
        LEFT JOIN electrical_equipment_voltages 
        ON electrical_equipment_voltages.id = electrical_lighting_timeclocks.voltage_id
        WHERE project_id = @projectId";
      OpenConnection();
      MySqlCommand command = new MySqlCommand(query, Connection);
      command.Parameters.AddWithValue("projectId", projectId);
      MySqlDataReader reader = command.ExecuteReader();
      while (reader.Read())
      {
        clocks.Add(
          new LightingTimeClock(
            GetSafeString(reader, "timeclock_id"),
            GetSafeString(reader, "name"),
            GetSafeString(reader, "bypass_switch_name"),
            GetSafeString(reader, "bypass_switch_location"),
            GetSafeString(reader, "adjacent_panel_id"),
            GetSafeInt(reader, "voltage").ToString()
          )
        );
      }
      CloseConnection();
      reader.Close();
      return clocks;
    }

    public string IdToVoltage(int voltageId)
    {
      string voltage = "0";
      switch (voltageId)
      {
        case (1):
          voltage = "115";
          break;
        case (2):
          voltage = "120";
          break;
        case (3):
          voltage = "208";
          break;
        case (4):
          voltage = "230";
          break;
        case (5):
          voltage = "240";
          break;
        case (6):
          voltage = "277";
          break;
        case (7):
          voltage = "460";
          break;
        case (8):
          voltage = "480";
          break;
      }
      return voltage;
    }

    public Dictionary<string, ObservableCollection<ElectricalKeyedNoteTable>> GetKeyedNoteTables(
      string projectId
    )
    {
      List<ElectricalKeyedNote> notes = new List<ElectricalKeyedNote>();
      List<ElectricalKeyedNoteTable> tables = new List<ElectricalKeyedNoteTable>();
      Dictionary<string, ObservableCollection<ElectricalKeyedNoteTable>> tablesDict =
        new Dictionary<string, ObservableCollection<ElectricalKeyedNoteTable>>();
      string query =
        @"
        SELECT 
        electrical_keyed_notes.id,
        electrical_keyed_notes.table_id,
        electrical_keyed_notes.date_created,
        electrical_keyed_notes.note
        FROM electrical_keyed_notes 
        WHERE project_id = @projectId
        order by date_created, table_id
        ";
      OpenConnection();
      MySqlCommand command = new MySqlCommand(query, Connection);
      command.Parameters.AddWithValue("projectId", projectId);
      MySqlDataReader reader = command.ExecuteReader();
      while (reader.Read())
      {
        notes.Add(
          new ElectricalKeyedNote
          {
            Id = GetSafeString(reader, "id"),
            TableId = GetSafeString(reader, "table_id"),
            DateCreated = reader.GetDateTime("date_created"),
            Note = GetSafeString(reader, "note"),
          }
        );
      }
      reader.Close();
      //Tables Query
      query =
        @"
        SELECT 
        electrical_keyed_note_tables.id,
        electrical_keyed_note_tables.sheet_id,
        electrical_keyed_note_tables.title
        FROM electrical_keyed_note_tables
        WHERE project_id = @projectId
        ";
      command = new MySqlCommand(query, Connection);
      command.Parameters.AddWithValue("projectId", projectId);
      reader = command.ExecuteReader();
      while (reader.Read())
      {
        tables.Add(
          new ElectricalKeyedNoteTable
          {
            Id = GetSafeString(reader, "id"),
            SheetId = GetSafeString(reader, "sheet_id"),
            Title = GetSafeString(reader, "title"),
          }
        );
      }
      reader.Close();

      //Add notes to tables
      foreach (var table in tables)
      {
        foreach (var note in notes)
        {
          if (note.TableId == table.Id)
          {
            table.KeyedNotes.Add(note);
          }
        }
        table.KeyedNotes = new BindingList<ElectricalKeyedNote>(
          table.KeyedNotes.OrderBy(n => n.DateCreated).ToList()
        );
      }
      foreach (var table in tables)
      {
        if (!tablesDict.ContainsKey(table.SheetId))
        {
          tablesDict.Add(table.SheetId, new ObservableCollection<ElectricalKeyedNoteTable>());
        }
        tablesDict[table.SheetId].Add(table);
      }

      CloseConnection();

      return tablesDict;
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

    public List<string> GetObjectIds(string projectId, string tableName)
    {
      List<string> objectIds = new List<string>();
      string query = $"SELECT id FROM {tableName} WHERE project_id = @projectId";
      OpenConnection();
      MySqlCommand command = new MySqlCommand(query, Connection);
      command.Parameters.AddWithValue("projectId", projectId);
      MySqlDataReader reader = command.ExecuteReader();
      while (reader.Read())
      {
        objectIds.Add(GetSafeString(reader, "id"));
      }
      reader.Close();
      CloseConnection();
      return objectIds;
    }

    public void UpdateKeyNotesTables(
      string projectId,
      Dictionary<string, ObservableCollection<ElectricalKeyedNoteTable>> tables
    )
    {
      DeleteObsoleteNotesAndTables(projectId, tables);

      string query =
        @"
        INSERT INTO electrical_keyed_note_tables (id, project_id, sheet_id, title)
        VALUES (@id, @projectId, @sheetId, @title)
        ON DUPLICATE KEY UPDATE
        sheet_id = @sheetId,
        title = @title
        ";
      OpenConnection();
      MySqlCommand command = new MySqlCommand(query, Connection);

      foreach (var kvp in tables)
      {
        foreach (var table in kvp.Value)
        {
          command.Parameters.AddWithValue("@projectId", projectId);
          command.Parameters.AddWithValue("@id", table.Id);
          command.Parameters.AddWithValue("@sheetId", table.SheetId);
          command.Parameters.AddWithValue("@title", table.Title);
          command.ExecuteNonQuery();
          command.Parameters.Clear();
        }
      }
      query =
        @"
        INSERT INTO electrical_keyed_notes (id, project_id, table_id, date_created, note)
        VALUES (@id, @projectId, @tableId, @dateCreated, @note)
        ON DUPLICATE KEY UPDATE
        date_created = @dateCreated,
        note = @note
        ";
      command = new MySqlCommand(query, Connection);

      foreach (var kvp in tables)
      {
        foreach (var table in kvp.Value)
        {
          foreach (var note in table.KeyedNotes)
          {
            command.Parameters.AddWithValue("@projectId", projectId);
            command.Parameters.AddWithValue("@id", note.Id);
            command.Parameters.AddWithValue("@tableId", note.TableId);
            command.Parameters.AddWithValue("@dateCreated", note.DateCreated);
            command.Parameters.AddWithValue("@note", note.Note);
            command.ExecuteNonQuery();
            command.Parameters.Clear();
          }
        }
      }
      CloseConnection();
    }

    public void DeleteObsoleteNotesAndTables(
      string projectId,
      Dictionary<string, ObservableCollection<ElectricalKeyedNoteTable>> tables
    )
    {
      // Retrieve current note and table IDs from the database
      List<string> currentNoteIds = GetObjectIds(projectId, "electrical_keyed_notes");
      List<string> currentTableIds = GetObjectIds(projectId, "electrical_keyed_note_tables");

      // Create sets of new note and table IDs from the dictionary
      HashSet<string> newNoteIds = new HashSet<string>();
      HashSet<string> newTableIds = new HashSet<string>();

      foreach (var kvp in tables)
      {
        foreach (var table in kvp.Value)
        {
          newTableIds.Add(table.Id);
          foreach (var note in table.KeyedNotes)
          {
            newNoteIds.Add(note.Id);
          }
        }
      }

      // Find obsolete notes and tables
      List<string> obsoleteNoteIds = currentNoteIds.Except(newNoteIds).ToList();
      List<string> obsoleteTableIds = currentTableIds.Except(newTableIds).ToList();

      // Delete obsolete notes
      if (obsoleteNoteIds.Count > 0)
      {
        string deleteNotesQuery =
          "DELETE FROM electrical_keyed_notes WHERE id IN ("
          + string.Join(",", obsoleteNoteIds.Select(id => $"'{id}'"))
          + ")";
        OpenConnection();
        MySqlCommand deleteNotesCommand = new MySqlCommand(deleteNotesQuery, Connection);
        deleteNotesCommand.ExecuteNonQuery();
        CloseConnection();
      }

      // Delete obsolete tables
      if (obsoleteTableIds.Count > 0)
      {
        string deleteTablesQuery =
          "DELETE FROM electrical_keyed_note_tables WHERE id IN ("
          + string.Join(",", obsoleteTableIds.Select(id => $"'{id}'"))
          + ")";
        OpenConnection();
        MySqlCommand deleteTablesCommand = new MySqlCommand(deleteTablesQuery, Connection);
        deleteTablesCommand.ExecuteNonQuery();
        CloseConnection();
      }
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
      command.Parameters.AddWithValue("@xLoc", equip.Location.X);
      command.Parameters.AddWithValue("@yLoc", equip.Location.Y);
      command.Parameters.AddWithValue("@parentDistance", equip.ParentDistance);
      command.Parameters.AddWithValue("@equipId", equip.Id);
      command.ExecuteNonQuery();
      CloseConnection();
    }

    public float ReadLightingFixtureWattage(string fixtureId)
    {
      float wattage = 0;
      string query =
        @"
        SELECT electrical_lighting.wattage
        FROM electrical_lighting
        WHERE electrical_lighting.id = @id
        ";
      this.OpenConnection();
      MySqlCommand command = new MySqlCommand(query, Connection);
      command.Parameters.AddWithValue("@id", fixtureId);
      MySqlDataReader reader = command.ExecuteReader();
      if (reader.Read())
      {
        wattage = GetSafeFloat(reader, "wattage");
      }
      reader.Close();

      return wattage;
    }

    public void UpdatePanelCircuitLoad(string panelId, string circuit, float va)
    {
      this.OpenConnection();
      string query =
        @"
        UPDATE electrical_equipment
        SET va = @va
        WHERE
        parent_id = @parentId
        AND
        circuit_no = @circuitNo
        AND
        circuit_half = @circuitHalf
        ";
      int circuitHalf = 0;
      if (circuit.Contains("A"))
      {
        circuitHalf = 1;
        circuit = circuit.Replace("A", "");
      }
      if (circuit.Contains("B"))
      {
        circuitHalf = 2;
        circuit = circuit.Replace("B", "");
      }

      int circuitNo = Int32.Parse(circuit);

      MySqlCommand command = new MySqlCommand(query, Connection);

      command.Parameters.AddWithValue("@parentId", panelId);
      command.Parameters.AddWithValue("@circuitNo", circuitNo);
      command.Parameters.AddWithValue("@circuitHalf", circuitHalf);
      command.Parameters.AddWithValue("@va", va);

      command.ExecuteNonQuery();

      CloseConnection();
    }

    public void InsertLightingEquipment(
      List<string> lightings,
      string panelId,
      int circuitNo,
      string projectId
    )
    {
      float newWattage = 0;
      string query =
        @"
        SELECT
        electrical_lighting.wattage
        FROM electrical_lighting
        WHERE electrical_lighting.id = @id";

      this.OpenConnection();
      foreach (var id in lightings)
      {
        MySqlCommand commandi = new MySqlCommand(query, Connection);
        commandi.Parameters.AddWithValue("@id", id);
        MySqlDataReader readeri = commandi.ExecuteReader();
        while (readeri.Read())
        {
          newWattage += GetSafeFloat(readeri, "wattage");
        }
        readeri.Close();
      }

      query =
        @"
      SELECT id, fla, voltage_id, category_id FROM electrical_equipment WHERE project_id = @projectId AND circuit_no = @circuitNo AND parent_id = @parentId
      ";

      MySqlCommand command = new MySqlCommand(query, Connection);

      command.Parameters.AddWithValue("@projectId", projectId);
      command.Parameters.AddWithValue("@parentId", panelId);
      command.Parameters.AddWithValue("@circuitNo", circuitNo);

      MySqlDataReader reader = command.ExecuteReader();

      float voltage = 120;
      if (reader.Read())
      {
        // there exists a circuit so just update the fla of it
        string id = GetSafeString(reader, "id");
        float fla = GetSafeFloat(reader, "fla");
        if (GetSafeInt(reader, "voltage_id") == 6)
        {
          voltage = 277;
        }
        reader.Close();
        query =
          @"
          UPDATE electrical_equipment SET
          fla = @fla,
          va = @va
          WHERE 
          id = @id
        ";

        MySqlCommand command2 = new MySqlCommand(query, Connection);
        command2.Parameters.AddWithValue("@id", id);
        command2.Parameters.AddWithValue(
          "@fla",
          fla + Math.Round(newWattage / voltage, 1, MidpointRounding.AwayFromZero)
        );
        command2.Parameters.AddWithValue("@va", fla * voltage + newWattage);
        command2.ExecuteNonQuery();
      }
      else
      {
        reader.Close();
        // new circuit
        query =
          @"INSERT INTO electrical_equipment (id, project_id, parent_id, description, category_id, voltage_id, 
        fla, is_three_phase, circuit_no, spec_sheet_from_client, aic_rating, color_code, connection_type_id, va, load_type) VALUES (@id, @projectId, @parentId, @description, @category, 
        @voltage, @fla, @isThreePhase, @circuit, @specFromClient, @aicRating, @colorCode, @connectionId, @va, @loadType)";

        MySqlCommand command2 = new MySqlCommand(query, Connection);
        command2.Parameters.AddWithValue("@id", Guid.NewGuid().ToString());
        command2.Parameters.AddWithValue("@projectId", projectId);
        command2.Parameters.AddWithValue("@parentId", panelId);
        command2.Parameters.AddWithValue("@description", "Lighting");
        command2.Parameters.AddWithValue("@category", 5);
        command2.Parameters.AddWithValue("@voltage", 1);
        command2.Parameters.AddWithValue(
          "@fla",
          Math.Round(newWattage / voltage, 1, MidpointRounding.AwayFromZero)
        );
        command2.Parameters.AddWithValue("@va", newWattage);
        command2.Parameters.AddWithValue("@isThreePhase", false);
        command2.Parameters.AddWithValue("@circuit", circuitNo);
        command2.Parameters.AddWithValue("@specFromClient", false);
        command2.Parameters.AddWithValue("@aicRating", 0);
        command2.Parameters.AddWithValue("@colorCode", "#FF00FF");
        command2.Parameters.AddWithValue("@connectionId", 1);
        command2.Parameters.AddWithValue("@loadType", 3);
        command2.ExecuteNonQuery();
      }

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

    public void updateLightingLocation(string lightingId, string locationId)
    {
      string query =
        @"
          UPDATE electrical_lighting
          SET
          location_id = @locationId
          WHERE id = @id
          ";
      OpenConnection();
      MySqlCommand command = new MySqlCommand(query, Connection);
      command.Parameters.AddWithValue("@locationId", locationId);
      command.Parameters.AddWithValue("@id", lightingId);
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
