// ================================================================
// MRP SYSTEM - DIAGRAMA DE CLASES
// Esquema DDL: Sistema MRP (PostgreSQL 16+)
// Ejecutar en EA: Scripting > JScript > Run
// ================================================================

function aviso(texto)
{
    var shell = new ActiveXObject("WScript.Shell");
    shell.Popup(texto, 0, "Progreso", 0);
}

function main()
{
    aviso("INICIO: Diagrama de Clases MRP...");
    try
    {
        // ---------- BUSQUEDA ----------
        function buscarPaquete(paquete, nombre) {
            if (paquete.Name == nombre) { return paquete; }
            for (var i = 0; i < paquete.Packages.Count; i++) {
                var e = buscarPaquete(paquete.Packages.GetAt(i), nombre);
                if (e != null) { return e; }
            }
            return null;
        }

        function obtenerPaquete(padre, nombre) {
            for (var i = 0; i < padre.Packages.Count; i++) {
                var sp = padre.Packages.GetAt(i);
                if (sp.Name == nombre) { return sp; }
            }
            return null;
        }

        function obtenerClase(padre, nombre) {
            for (var i = 0; i < padre.Elements.Count; i++) {
                var el = padre.Elements.GetAt(i);
                if (el.Name == nombre && el.Type == "Class") { return el; }
            }
            return null;
        }

        function obtenerDiagrama(paquete, nombre) {
            for (var i = 0; i < paquete.Diagrams.Count; i++) {
                var d = paquete.Diagrams.GetAt(i);
                if (d.Name == nombre) { return d; }
            }
            return null;
        }

        function yaEstaEnDiagrama(diagram, elementID) {
            for (var i = 0; i < diagram.DiagramObjects.Count; i++) {
                if (diagram.DiagramObjects.GetAt(i).ElementID == elementID) return true;
            }
            return false;
        }

        function posEnDiagrama(diagram, elemento, l, t, r, b) {
            if (yaEstaEnDiagrama(diagram, elemento.ElementID)) { return; }
            var obj = diagram.DiagramObjects.AddNew("l="+l+";r="+r+";t="+t+";b="+b+";", "");
            obj.ElementID = elemento.ElementID;
            obj.Update();
        }

        function existeAsociacion(elemento, supplierID) {
            for (var i = 0; i < elemento.Connectors.Count; i++) {
                var c = elemento.Connectors.GetAt(i);
                if (c.Type == "Association" && c.SupplierID == supplierID) return true;
            }
            return false;
        }

        function crearAtributos(clase, attrs) {
            // attrs = array de [nombre, tipo, longitud, esPK, esFK]
            for (var i = 0; i < attrs.length; i++) {
                var a = attrs[i];
                var attrName = a[0];
                var attrType = a[1];
                var attrLen = a[2] || "";
                var isPK = a[3] || false;
                var isFK = a[4] || false;
                var prefix = "";
                if (isPK) prefix = "{pk} ";
                if (isFK) prefix += "{fk} ";
                var typeName = attrType;
                if (attrLen !== "") typeName += "[" + attrLen + "]";
                var attr = clase.Attributes.AddNew(attrName, typeName);
                attr.Notes = (isPK ? "PK" : "") + (isFK ? " | FK" : "");
                attr.Update();
            }
            clase.Attributes.Refresh();
        }

        function crearClase(padre, nombre) {
            var el = obtenerClase(padre, nombre);
            if (el == null) {
                el = padre.Elements.AddNew(nombre, "Class");
                el.Update();
                padre.Elements.Refresh();
            }
            return el;
        }

        function linkAsociacion(diagram, origen, destino, nombre, multiplicidadO, multiplicidadD) {
            if (existeAsociacion(origen, destino.ElementID)) { return; }
            var conn = origen.Connectors.AddNew("", "Association");
            conn.ClientID = origen.ElementID;
            conn.SupplierID = destino.ElementID;
            if (nombre) { conn.Name = nombre; }
            conn.Update();
            try { conn.SourceEnd.Multiplicity = multiplicidadO; } catch(e) {}
            try { conn.DestEnd.Multiplicity = multiplicidadD; } catch(e) {}
            conn.Update();
            origen.Connectors.Refresh();
            var dl = diagram.DiagramLinks.AddNew("", "");
            dl.ConnectorID = conn.ConnectorID;
            dl.Update();
        }

        // ---------- PAQUETE ANCLA ----------
        var raizModelo = Repository.Models.GetAt(0);
        var paqMRP = buscarPaquete(raizModelo, "MRP System");
        if (paqMRP == null) {
            paqMRP = raizModelo.Packages.AddNew("MRP System", "");
            paqMRP.Update();
            raizModelo.Packages.Refresh();
        }
        aviso("Paquete ancla: '" + paqMRP.Name + "'");

        // ---------- DIAGRAMA ----------
        var nombreDiagrama = "MRP - Diagrama de Clases";
        var diagram = obtenerDiagrama(paqMRP, nombreDiagrama);
        if (diagram == null) {
            diagram = paqMRP.Diagrams.AddNew(nombreDiagrama, "Class");
            diagram.Update();
            paqMRP.Diagrams.Refresh();
        }
        aviso("Diagrama listo: '" + diagram.Name + "'");

        // =============================================
        // CREAR PAQUETES Y CLASES POR GRUPO
        // =============================================

        // ---- 1. SEGURIDAD Y USUARIOS ----
        var paqSeg = obtenerPaquete(paqMRP, "Seguridad y Usuarios");
        if (paqSeg == null) { paqSeg = paqMRP.Packages.AddNew("Seguridad y Usuarios", ""); paqSeg.Update(); paqMRP.Packages.Refresh(); }

        var clsRoles = crearClase(paqSeg, "Roles");
        crearAtributos(clsRoles, [
            ["id", "SERIAL", "", true, false],
            ["name", "VARCHAR(50)", 50, false, false],
            ["description", "VARCHAR(200)", 200, false, false],
            ["created_at", "TIMESTAMP", "", false, false]
        ]);

        var clsUsers = crearClase(paqSeg, "Users");
        crearAtributos(clsUsers, [
            ["id", "BIGSERIAL", "", true, false],
            ["uuid", "UUID", "", false, false],
            ["username", "VARCHAR(80)", 80, false, false],
            ["email", "VARCHAR(255)", 255, false, false],
            ["password_hash", "VARCHAR(255)", 255, false, false],
            ["first_name", "VARCHAR(100)", 100, false, false],
            ["last_name", "VARCHAR(100)", 100, false, false],
            ["is_active", "BOOLEAN", "", false, false],
            ["created_at", "TIMESTAMP", "", false, false],
            ["updated_at", "TIMESTAMP", "", false, false]
        ]);

        var clsUserRoles = crearClase(paqSeg, "UserRoles");
        crearAtributos(clsUserRoles, [
            ["user_id", "BIGINT", "", false, true],
            ["role_id", "INT", "", false, true],
            ["assigned_at", "TIMESTAMP", "", false, false]
        ]);

        // ---- 2. PROVEEDORES Y ALMACENES ----
        var paqProvAlm = obtenerPaquete(paqMRP, "Proveedores y Almacenes");
        if (paqProvAlm == null) { paqProvAlm = paqMRP.Packages.AddNew("Proveedores y Almacenes", ""); paqProvAlm.Update(); paqMRP.Packages.Refresh(); }

        var clsSuppliers = crearClase(paqProvAlm, "Suppliers");
        crearAtributos(clsSuppliers, [
            ["id", "BIGSERIAL", "", true, false],
            ["uuid", "UUID", "", false, false],
            ["code", "VARCHAR(50)", 50, false, false],
            ["name", "VARCHAR(180)", 180, false, false],
            ["legal_name", "VARCHAR(200)", 200, false, false],
            ["tax_id", "VARCHAR(50)", 50, false, false],
            ["email", "VARCHAR(255)", 255, false, false],
            ["phone", "VARCHAR(30)", 30, false, false],
            ["address", "TEXT", "", false, false],
            ["contact_name", "VARCHAR(150)", 150, false, false],
            ["is_active", "BOOLEAN", "", false, false],
            ["created_at", "TIMESTAMP", "", false, false],
            ["updated_at", "TIMESTAMP", "", false, false]
        ]);

        var clsWarehouses = crearClase(paqProvAlm, "Warehouses");
        crearAtributos(clsWarehouses, [
            ["id", "BIGSERIAL", "", true, false],
            ["uuid", "UUID", "", false, false],
            ["code", "VARCHAR(50)", 50, false, false],
            ["name", "VARCHAR(150)", 150, false, false],
            ["address", "TEXT", "", false, false],
            ["is_active", "BOOLEAN", "", false, false],
            ["created_at", "TIMESTAMP", "", false, false],
            ["updated_at", "TIMESTAMP", "", false, false]
        ]);

        // ---- 3. ARTICULOS Y STOCK ----
        var paqArtStock = obtenerPaquete(paqMRP, "Articulos y Stock");
        if (paqArtStock == null) { paqArtStock = paqMRP.Packages.AddNew("Articulos y Stock", ""); paqArtStock.Update(); paqMRP.Packages.Refresh(); }

        var clsItems = crearClase(paqArtStock, "Items");
        crearAtributos(clsItems, [
            ["id", "BIGSERIAL", "", true, false],
            ["uuid", "UUID", "", false, false],
            ["code", "VARCHAR(60)", 60, false, false],
            ["name", "VARCHAR(200)", 200, false, false],
            ["description", "TEXT", "", false, false],
            ["item_type", "VARCHAR(30)", 30, false, false],
            ["uom", "VARCHAR(20)", 20, false, false],
            ["safety_stock", "DECIMAL(14,4)", "14,4", false, false],
            ["lead_time_days", "INT", "", false, false],
            ["standard_cost", "DECIMAL(14,4)", "14,4", false, false],
            ["preferred_supplier_id", "BIGINT", "", false, true],
            ["is_active", "BOOLEAN", "", false, false],
            ["created_at", "TIMESTAMP", "", false, false],
            ["updated_at", "TIMESTAMP", "", false, false]
        ]);

        var clsBatches = crearClase(paqArtStock, "Batches");
        crearAtributos(clsBatches, [
            ["id", "BIGSERIAL", "", true, false],
            ["uuid", "UUID", "", false, false],
            ["batch_number", "VARCHAR(80)", 80, false, false],
            ["supplier_id", "BIGINT", "", false, true],
            ["warehouse_id", "BIGINT", "", false, true],
            ["received_date", "TIMESTAMP", "", false, false],
            ["status", "VARCHAR(20)", 20, false, false],
            ["notes", "TEXT", "", false, false],
            ["created_by_id", "BIGINT", "", false, true],
            ["created_at", "TIMESTAMP", "", false, false],
            ["updated_at", "TIMESTAMP", "", false, false]
        ]);

        var clsBatchDetails = crearClase(paqArtStock, "BatchDetails");
        crearAtributos(clsBatchDetails, [
            ["id", "BIGSERIAL", "", true, false],
            ["batch_id", "BIGINT", "", false, true],
            ["item_id", "BIGINT", "", false, true],
            ["quantity_received", "DECIMAL(14,4)", "14,4", false, false],
            ["quantity_available", "DECIMAL(14,4)", "14,4", false, false],
            ["unit_cost", "DECIMAL(14,4)", "14,4", false, false],
            ["expiration_date", "DATE", "", false, false],
            ["created_at", "TIMESTAMP", "", false, false],
            ["updated_at", "TIMESTAMP", "", false, false]
        ]);

        var clsInventoryStocks = crearClase(paqArtStock, "InventoryStocks");
        crearAtributos(clsInventoryStocks, [
            ["id", "BIGSERIAL", "", true, false],
            ["warehouse_id", "BIGINT", "", false, true],
            ["item_id", "BIGINT", "", false, true],
            ["physical_stock", "DECIMAL(14,4)", "14,4", false, false],
            ["allocated_stock", "DECIMAL(14,4)", "14,4", false, false],
            ["available_stock", "DECIMAL(14,4)", "14,4", false, false],
            ["updated_at", "TIMESTAMP", "", false, false]
        ]);

        var clsInventoryMovements = crearClase(paqArtStock, "InventoryMovements");
        crearAtributos(clsInventoryMovements, [
            ["id", "BIGSERIAL", "", true, false],
            ["uuid", "UUID", "", false, false],
            ["movement_type", "VARCHAR(35)", 35, false, false],
            ["warehouse_id", "BIGINT", "", false, true],
            ["item_id", "BIGINT", "", false, true],
            ["batch_id", "BIGINT", "", false, false],
            ["quantity", "DECIMAL(14,4)", "14,4", false, false],
            ["previous_stock", "DECIMAL(14,4)", "14,4", false, false],
            ["new_stock", "DECIMAL(14,4)", "14,4", false, false],
            ["reason", "VARCHAR(255)", 255, false, false],
            ["reference_document", "VARCHAR(120)", 120, false, false],
            ["registered_by_id", "BIGINT", "", false, true],
            ["created_at", "TIMESTAMP", "", false, false]
        ]);

        // ---- 4. LISTA DE MATERIALES (BOM) ----
        var paqBOM = obtenerPaquete(paqMRP, "Lista de Materiales");
        if (paqBOM == null) { paqBOM = paqMRP.Packages.AddNew("Lista de Materiales", ""); paqBOM.Update(); paqMRP.Packages.Refresh(); }

        var clsBillsOfMaterials = crearClase(paqBOM, "BillsOfMaterials");
        crearAtributos(clsBillsOfMaterials, [
            ["id", "BIGSERIAL", "", true, false],
            ["uuid", "UUID", "", false, false],
            ["code", "VARCHAR(60)", 60, false, false],
            ["name", "VARCHAR(200)", 200, false, false],
            ["finished_item_id", "BIGINT", "", false, true],
            ["version", "VARCHAR(20)", 20, false, false],
            ["base_quantity", "DECIMAL(14,4)", "14,4", false, false],
            ["is_active", "BOOLEAN", "", false, false],
            ["notes", "TEXT", "", false, false],
            ["created_by_id", "BIGINT", "", false, true],
            ["created_at", "TIMESTAMP", "", false, false],
            ["updated_at", "TIMESTAMP", "", false, false]
        ]);

        var clsBomItems = crearClase(paqBOM, "BomItems");
        crearAtributos(clsBomItems, [
            ["id", "BIGSERIAL", "", true, false],
            ["bom_id", "BIGINT", "", false, true],
            ["component_item_id", "BIGINT", "", false, true],
            ["quantity_required", "DECIMAL(14,4)", "14,4", false, false],
            ["scrap_percentage", "DECIMAL(5,2)", "5,2", false, false],
            ["notes", "TEXT", "", false, false],
            ["created_at", "TIMESTAMP", "", false, false]
        ]);

        // ---- 5. MOTOR MRP ----
        var paqMRPExec = obtenerPaquete(paqMRP, "Motor MRP");
        if (paqMRPExec == null) { paqMRPExec = paqMRP.Packages.AddNew("Motor MRP", ""); paqMRPExec.Update(); paqMRPExec.Packages.Refresh(); }

        var clsMRPRuns = crearClase(paqMRPExec, "MRPRuns");
        crearAtributos(clsMRPRuns, [
            ["id", "BIGSERIAL", "", true, false],
            ["uuid", "UUID", "", false, false],
            ["run_code", "VARCHAR(80)", 80, false, false],
            ["finished_item_id", "BIGINT", "", false, true],
            ["bom_id", "BIGINT", "", false, true],
            ["planned_quantity", "DECIMAL(14,4)", "14,4", false, false],
            ["target_date", "DATE", "", false, false],
            ["status", "VARCHAR(30)", 30, false, false],
            ["executed_by_id", "BIGINT", "", false, true],
            ["created_at", "TIMESTAMP", "", false, false]
        ]);

        var clsMRPRunItems = crearClase(paqMRPExec, "MRPRunItems");
        crearAtributos(clsMRPRunItems, [
            ["id", "BIGSERIAL", "", true, false],
            ["mrp_run_id", "BIGINT", "", false, true],
            ["component_item_id", "BIGINT", "", false, true],
            ["gross_requirement", "DECIMAL(14,4)", "14,4", false, false],
            ["current_physical_stock", "DECIMAL(14,4)", "14,4", false, false],
            ["current_allocated_stock", "DECIMAL(14,4)", "14,4", false, false],
            ["current_available_stock", "DECIMAL(14,4)", "14,4", false, false],
            ["safety_stock", "DECIMAL(14,4)", "14,4", false, false],
            ["net_requirement", "DECIMAL(14,4)", "14,4", false, false],
            ["supplier_id", "BIGINT", "", false, false],
            ["lead_time_days", "INT", "", false, false],
            ["suggested_order_date", "DATE", "", false, false],
            ["estimated_unit_cost", "DECIMAL(14,4)", "14,4", false, false],
            ["total_estimated_cost", "DECIMAL(14,4)", "14,4", false, false],
            ["created_at", "TIMESTAMP", "", false, false]
        ]);

        // =============================================
        // CREAR ASOCIACIONES ENTRE CLASES
        // =============================================
        aviso("Creando asociaciones...");

        // Users <-> Roles (many-to-many)
        linkAsociacion(diagram, clsUsers, clsUserRoles, "", "1", "*");
        linkAsociacion(diagram, clsRoles, clsUserRoles, "", "1", "*");

        // Suppliers -> Items (preferred supplier)
        linkAsociacion(diagram, clsSuppliers, clsItems, "preferred", "1", "*");

        // Suppliers -> Batches
        linkAsociacion(diagram, clsSuppliers, clsBatches, "", "1", "*");

        // Warehouses -> Batches
        linkAsociacion(diagram, clsWarehouses, clsBatches, "", "1", "*");

        // Users -> Batches (created_by)
        linkAsociacion(diagram, clsUsers, clsBatches, "created_by", "1", "*");

        // Batches -> BatchDetails
        linkAsociacion(diagram, clsBatches, clsBatchDetails, "", "1", "*");

        // Items -> BatchDetails
        linkAsociacion(diagram, clsItems, clsBatchDetails, "", "1", "*");

        // Warehouses -> InventoryStocks
        linkAsociacion(diagram, clsWarehouses, clsInventoryStocks, "", "1", "*");

        // Items -> InventoryStocks
        linkAsociacion(diagram, clsItems, clsInventoryStocks, "", "1", "*");

        // Warehouses -> InventoryMovements
        linkAsociacion(diagram, clsWarehouses, clsInventoryMovements, "", "1", "*");

        // Items -> InventoryMovements
        linkAsociacion(diagram, clsItems, clsInventoryMovements, "", "1", "*");

        // Batches -> InventoryMovements
        linkAsociacion(diagram, clsBatches, clsInventoryMovements, "", "0..1", "*");

        // Users -> InventoryMovements
        linkAsociacion(diagram, clsUsers, clsInventoryMovements, "registered_by", "1", "*");

        // Items -> BillsOfMaterials (finished_item)
        linkAsociacion(diagram, clsItems, clsBillsOfMaterials, "finished", "1", "*");

        // Users -> BillsOfMaterials (created_by)
        linkAsociacion(diagram, clsUsers, clsBillsOfMaterials, "created_by", "1", "*");

        // BillsOfMaterials -> BomItems
        linkAsociacion(diagram, clsBillsOfMaterials, clsBomItems, "", "1", "*");

        // Items -> BomItems (component)
        linkAsociacion(diagram, clsItems, clsBomItems, "", "1", "*");

        // Items -> MRPRuns (finished_item)
        linkAsociacion(diagram, clsItems, clsMRPRuns, "finished", "1", "*");

        // BillsOfMaterials -> MRPRuns
        linkAsociacion(diagram, clsBillsOfMaterials, clsMRPRuns, "", "1", "*");

        // Users -> MRPRuns (executed_by)
        linkAsociacion(diagram, clsUsers, clsMRPRuns, "executed_by", "1", "*");

        // MRPRuns -> MRPRunItems
        linkAsociacion(diagram, clsMRPRuns, clsMRPRunItems, "", "1", "*");

        // Items -> MRPRunItems
        linkAsociacion(diagram, clsItems, clsMRPRunItems, "", "1", "*");

        // Suppliers -> MRPRunItems
        linkAsociacion(diagram, clsSuppliers, clsMRPRunItems, "", "0..1", "*");

        // =============================================
        // POSICIONAR CLASES EN EL DIAGRAMA
        // =============================================
        aviso("Posicionando clases en el diagrama...");

        // FILA 1: Seguridad y Usuarios (arriba izquierda)
        posEnDiagrama(diagram, clsRoles, 30, 30, 230, 200);
        posEnDiagrama(diagram, clsUsers, 30, 230, 230, 500);
        posEnDiagrama(diagram, clsUserRoles, 30, 530, 230, 680);

        // FILA 2: Proveedores y Almacenes (arriba centro)
        posEnDiagrama(diagram, clsSuppliers, 320, 30, 520, 260);
        posEnDiagrama(diagram, clsWarehouses, 320, 290, 520, 490);

        // FILA 3: Articulos y Stock (arriba derecha)
        posEnDiagrama(diagram, clsItems, 610, 30, 810, 350);
        posEnDiagrama(diagram, clsBatches, 610, 380, 810, 550);
        posEnDiagrama(diagram, clsBatchDetails, 610, 580, 810, 730);
        posEnDiagrama(diagram, clsInventoryStocks, 610, 760, 810, 910);
        posEnDiagrama(diagram, clsInventoryMovements, 610, 940, 810, 1090);

        // FILA 4: Lista de Materiales (abajo izquierda)
        posEnDiagrama(diagram, clsBillsOfMaterials, 30, 720, 230, 920);
        posEnDiagrama(diagram, clsBomItems, 30, 950, 230, 1100);

        // FILA 5: Motor MRP (abajo derecha)
        posEnDiagrama(diagram, clsMRPRuns, 320, 720, 520, 890);
        posEnDiagrama(diagram, clsMRPRunItems, 320, 920, 520, 1100);

        diagram.Update();

        aviso("Diagrama de Clases MRP completado!\n\n"
            + "Clases creadas: 14\n"
            + "Asociaciones creadas: 26\n\n"
            + "Abriendo diagrama...");

        diagram.Update();
        Repository.OpenDiagram(diagram.DiagramID);
    }
    catch (e)
    {
        var shell = new ActiveXObject("WScript.Shell");
        shell.Popup("ERROR: " + e.description, 0, "Error", 0);
    }
}

main();
