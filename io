<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Zetas Pizzería - Caja, Pedidos, Clientes y Fidelización</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Caveat:wght@600;700&family=Montserrat:wght@400;600;700;900&display=swap" rel="stylesheet">
  <!-- Librería SheetJS para exportar a Microsoft Excel (.xlsx) nativo -->
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <style>
    :root {
      --bg-dark: #121316;
      --bg-card: #1c1e24;
      --bg-input: #272a33;
      --orange-accent: #ff6a00;
      --orange-hover: #ff852e;
      --chalk-white: #f5f6f8;
      --chalk-muted: #9ba1b0;
      --green-cash: #22c55e;
      --red-alert: #ef4444;
      --gold-vip: #f59e0b;
      --border-chalk: rgba(255, 255, 255, 0.12);
      --font-chalk: 'Caveat', cursive;
      --font-ui: 'Montserrat', sans-serif;
    }

    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      background-color: var(--bg-dark);
      color: var(--chalk-white);
      font-family: var(--font-ui);
      min-height: 100vh;
      padding-bottom: 50px;
    }

    body::before {
      content: "";
      position: fixed;
      inset: 0;
      background-image: radial-gradient(rgba(255, 255, 255, 0.03) 1px, transparent 0);
      background-size: 24px 24px;
      pointer-events: none;
      z-index: 0;
    }

    .container {
      position: relative;
      z-index: 1;
      max-width: 1280px;
      margin: 0 auto;
      padding: 16px;
    }

    header {
      text-align: center;
      margin-bottom: 20px;
      border-bottom: 2px dashed var(--border-chalk);
      padding-bottom: 16px;
    }

    .logo-title {
      font-family: var(--font-chalk);
      font-size: 3.6rem;
      font-weight: 700;
      color: #fff;
      text-shadow: 2px 2px 0px var(--orange-accent);
      line-height: 1;
    }

    .subtitle-badge {
      display: inline-block;
      background: var(--orange-accent);
      color: #000;
      font-weight: 900;
      font-size: 0.8rem;
      letter-spacing: 1px;
      text-transform: uppercase;
      padding: 4px 12px;
      border-radius: 4px;
      margin-top: 6px;
    }

    /* TABS */
    .tabs-nav {
      display: flex;
      gap: 8px;
      margin-bottom: 20px;
      border-bottom: 1px solid var(--border-chalk);
      overflow-x: auto;
      padding-bottom: 8px;
    }

    .tab-btn {
      background: var(--bg-card);
      border: 1px solid var(--border-chalk);
      color: var(--chalk-muted);
      padding: 10px 16px;
      border-radius: 8px;
      font-weight: 700;
      font-size: 0.88rem;
      cursor: pointer;
      display: flex;
      align-items: center;
      gap: 6px;
      white-space: nowrap;
      transition: all 0.2s ease;
    }

    .tab-btn:hover { background: var(--bg-input); color: #fff; }
    .tab-btn.active {
      background: var(--orange-accent);
      color: #000;
      border-color: var(--orange-accent);
      box-shadow: 0 4px 12px rgba(255, 106, 0, 0.35);
    }

    .tab-panel { display: none; }
    .tab-panel.active { display: block; animation: fadeIn 0.25s ease; }

    @keyframes fadeIn {
      from { opacity: 0; transform: translateY(6px); }
      to { opacity: 1; transform: translateY(0); }
    }

    .card {
      background: var(--bg-card);
      border: 1px solid var(--border-chalk);
      border-radius: 12px;
      padding: 20px;
      margin-bottom: 20px;
      box-shadow: 0 4px 14px rgba(0, 0, 0, 0.35);
    }

    .card-title {
      font-family: var(--font-chalk);
      font-size: 1.8rem;
      color: var(--orange-accent);
      margin-bottom: 14px;
      display: flex;
      align-items: center;
      justify-content: space-between;
    }

    .grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; }
    .grid-3 { display: grid; grid-template-columns: repeat(3, 1fr); gap: 14px; }
    .grid-4 { display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; }

    @media (max-width: 900px) {
      .grid-2, .grid-3, .grid-4 { grid-template-columns: 1fr; }
    }

    .form-group { margin-bottom: 14px; }
    label {
      display: block;
      font-size: 0.78rem;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      color: var(--chalk-muted);
      margin-bottom: 5px;
    }

    input, select, textarea {
      width: 100%;
      background: var(--bg-input);
      border: 1px solid var(--border-chalk);
      border-radius: 8px;
      padding: 10px 12px;
      color: #fff;
      font-family: var(--font-ui);
      font-size: 0.9rem;
      outline: none;
    }

    input:focus, select:focus, textarea:focus {
      border-color: var(--orange-accent);
    }

    .section-subtitle {
      font-weight: 700;
      color: #fff;
      font-size: 0.92rem;
      margin: 12px 0 8px 0;
      display: flex;
      align-items: center;
      justify-content: space-between;
    }

    .counter-pill {
      background: var(--bg-input);
      color: var(--orange-accent);
      padding: 2px 8px;
      border-radius: 10px;
      font-size: 0.8rem;
      border: 1px solid var(--orange-accent);
    }

    .ingredients-grid {
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(125px, 1fr));
      gap: 6px;
      margin-bottom: 14px;
    }

    .choice-chip { position: relative; }
    .choice-chip input { position: absolute; opacity: 0; width: 0; height: 0; }
    .chip-label {
      display: block;
      padding: 8px 6px;
      background: var(--bg-input);
      border: 1px solid var(--border-chalk);
      border-radius: 6px;
      font-size: 0.8rem;
      font-weight: 600;
      text-align: center;
      cursor: pointer;
      user-select: none;
      transition: 0.15s ease;
    }

    .choice-chip input:checked + .chip-label {
      background: var(--orange-accent);
      color: #000;
      font-weight: 800;
      border-color: var(--orange-accent);
    }

    .choice-chip.disabled { opacity: 0.35; cursor: not-allowed; }

    .additionals-grid {
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(155px, 1fr));
      gap: 6px;
      margin-bottom: 14px;
    }

    .additionals-grid .choice-chip input:checked + .chip-label {
      background: #ffb74d;
      color: #000;
      border-color: #ffb74d;
    }

    .products-item {
      display: flex;
      align-items: center;
      justify-content: space-between;
      background: var(--bg-input);
      padding: 8px 12px;
      border-radius: 6px;
      border: 1px solid var(--border-chalk);
    }

    .qty-controls {
      display: flex;
      align-items: center;
      gap: 8px;
    }

    .qty-btn {
      background: var(--bg-card);
      border: 1px solid var(--border-chalk);
      color: #fff;
      width: 28px;
      height: 28px;
      border-radius: 4px;
      font-weight: 800;
      cursor: pointer;
    }

    .btn {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      gap: 6px;
      padding: 10px 18px;
      border-radius: 8px;
      font-weight: 800;
      font-size: 0.9rem;
      cursor: pointer;
      border: none;
      transition: all 0.2s ease;
    }

    .btn-orange { background: var(--orange-accent); color: #000; }
    .btn-orange:hover { background: var(--orange-hover); }
    .btn-green { background: var(--green-cash); color: #000; }
    .btn-green:hover { filter: brightness(1.1); }
    .btn-danger { background: rgba(239, 68, 68, 0.15); border: 1px solid var(--red-alert); color: var(--red-alert); }
    .btn-danger:hover { background: var(--red-alert); color: #fff; }
    .btn-ghost { background: transparent; border: 1px solid var(--border-chalk); color: var(--chalk-muted); }
    .btn-ghost:hover { color: #fff; border-color: #fff; }

    .table-container {
      overflow-x: auto;
      border-radius: 8px;
      border: 1px solid var(--border-chalk);
    }

    table { width: 100%; border-collapse: collapse; font-size: 0.82rem; text-align: left; }
    th {
      background: var(--bg-input);
      color: var(--chalk-muted);
      text-transform: uppercase;
      font-size: 0.72rem;
      font-weight: 700;
      padding: 10px 12px;
      border-bottom: 1px solid var(--border-chalk);
      white-space: nowrap;
    }

    td { padding: 10px 12px; border-bottom: 1px solid var(--border-chalk); white-space: nowrap; }
    tr:hover td { background: rgba(255, 255, 255, 0.02); }

    .tag { padding: 2px 6px; border-radius: 4px; font-size: 0.72rem; font-weight: 700; }
    .tag-points { background: #f59e0b; color: #000; font-weight: 800; }

    /* DASHBOARD */
    .kpi-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
      gap: 12px;
      margin-bottom: 20px;
    }

    .kpi-card {
      background: var(--bg-card);
      border: 1px solid var(--border-chalk);
      border-radius: 10px;
      padding: 16px;
      border-left: 4px solid var(--orange-accent);
    }

    .kpi-title { font-size: 0.75rem; text-transform: uppercase; color: var(--chalk-muted); font-weight: 700; }
    .kpi-value { font-size: 1.6rem; font-weight: 900; color: #fff; margin: 4px 0; }
    .kpi-sub { font-size: 0.72rem; color: var(--chalk-muted); }

    .bar-row { margin-bottom: 10px; }
    .bar-header { display: flex; justify-content: space-between; font-size: 0.8rem; margin-bottom: 3px; }
    .bar-track { height: 8px; background: var(--bg-input); border-radius: 4px; overflow: hidden; }
    .bar-fill { height: 100%; background: var(--orange-accent); border-radius: 4px; }
  </style>
</head>
<body>

  <div class="container">
    <header>
      <div class="logo-title">Zetas</div>
      <div style="font-family: var(--font-chalk); font-size: 1.6rem; color: #fff;">pizzería</div>
      <div class="subtitle-badge">Sistema Integral: Ventas, Clientes, Descuentos & Fidelización</div>
    </header>

    <div class="tabs-nav">
      <button class="tab-btn active" onclick="switchTab('ventas')">🍕 1. Pedido / Venta</button>
      <button class="tab-btn" onclick="switchTab('clientes')">⭐ 2. Clientes & Puntos</button>
      <button class="tab-btn" onclick="switchTab('gastos')">💸 3. Gastos Caja Chica</button>
      <button class="tab-btn" onclick="switchTab('registros')">📋 4. Historial & Excel</button>
      <button class="tab-btn" onclick="switchTab('dashboard')">📊 5. Dashboard de Decisiones</button>
    </div>

    <!-- PANEL 1: VENTAS -->
    <div id="tab-ventas" class="tab-panel active">
      <div class="card">
        <div class="card-title">
          <span>Nuevo Pedido</span>
          <span style="font-size: 0.9rem; color: var(--chalk-muted); font-family: var(--font-ui);">Base Pizza: $9.000</span>
        </div>

        <form id="orderForm" onsubmit="saveOrder(event)">
          <!-- DATOS CLIENTE -->
          <div style="background: rgba(255, 106, 0, 0.08); border: 1px dashed var(--orange-accent); border-radius: 8px; padding: 14px; margin-bottom: 16px;">
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px;">
              <span style="font-weight: 800; font-size: 0.85rem; text-transform: uppercase; color: var(--orange-accent);">👤 Datos del Cliente & Despacho</span>
              <button type="button" class="btn btn-ghost" style="padding: 4px 8px; font-size: 0.75rem;" onclick="openQuickClientSearch()">🔍 Buscar Registrado</button>
            </div>
            <div class="grid-3">
              <div class="form-group" style="margin-bottom:0;">
                <label>Nombre Cliente</label>
                <input type="text" id="clientName" placeholder="Ej: Marcela Soto" required oninput="checkClientMatch()">
              </div>
              <div class="form-group" style="margin-bottom:0;">
                <label>Teléfono / WhatsApp</label>
                <input type="tel" id="clientPhone" placeholder="Ej: +56 9 8765 4321" required>
              </div>
              <div class="form-group" style="margin-bottom:0;">
                <label>Dirección de Despacho</label>
                <input type="text" id="clientAddress" placeholder="Calle, N°, Pasaje, Depto" required>
              </div>
            </div>
            <div id="clientPointsNotice" style="margin-top: 8px; font-size: 0.8rem; color: var(--gold-vip); font-weight: 700; display: none;">
              ⭐ Cliente frecuente: <span id="clientPointsCount">0</span> puntos acumulados.
            </div>
          </div>

          <!-- PARÁMETROS OPERATIVOS -->
          <div class="grid-3">
            <div class="form-group">
              <label>Turno</label>
              <select id="saleTurno">
                <option value="Noche">Cena / Noche</option>
                <option value="Almuerzo">Almuerzo</option>
              </select>
            </div>
            <div class="form-group">
              <label>Canal de Venta</label>
              <select id="saleChannel" onchange="calcTotalOrder()">
                <option value="Delivery Propio">Delivery Propio</option>
                <option value="Retiro">Retiro en Local</option>
                <option value="Local">Consumo en Mostrador / Local</option>
                <option value="App">App Delivery (PedidosYa/Uber)</option>
              </select>
            </div>
            <div class="form-group">
              <label>Método de Pago</label>
              <select id="salePayment">
                <option value="Efectivo">Efectivo</option>
                <option value="Débito / POS">Débito / Tarjeta POS</option>
                <option value="Transferencia">Transferencia</option>
                <option value="Pago en App">Pago Digital en App</option>
              </select>
            </div>
          </div>

          <!-- 18 INGREDIENTES OFICIALES -->
          <div class="section-subtitle">
            <span>1. Ingredientes Base de Pizza (Elige hasta 5)</span>
            <span class="counter-pill" id="baseCounter">0 / 5 seleccionados</span>
          </div>
          <div class="ingredients-grid" id="baseIngredientsContainer"></div>

          <!-- ADICIONALES (+$2.000) -->
          <div class="section-subtitle">
            <span>2. Adicionales Premium (+$2.000 c/u)</span>
            <span class="counter-pill" id="extraCounter" style="color: #ffb74d; border-color: #ffb74d;">0 adicionales</span>
          </div>
          <div class="additionals-grid" id="extraIngredientsContainer"></div>

          <!-- BEBIDAS, MONSTER, JUGOS Y ACOMPAÑAMIENTOS -->
          <div class="section-subtitle">
            <span>3. Bebidas, Energéticas Monster & Acompañamientos</span>
            <span style="font-size: 0.8rem; color: var(--chalk-muted); font-weight: normal;">Cantidades</span>
          </div>
          <div class="grid-3" id="beveragesContainer"></div>

          <!-- DESCUENTOS, REGALOS Y TOTALES -->
          <div style="background: var(--bg-input); border-radius: 8px; padding: 14px; margin: 16px 0;">
            <div class="grid-4">
              <div class="form-group" style="margin-bottom:0;">
                <label>Delivery ($)</label>
                <input type="number" id="deliveryFee" value="1500" step="500" oninput="calcTotalOrder()">
              </div>
              <div class="form-group" style="margin-bottom:0;">
                <label>Tipo Descuento / Beneficio</label>
                <select id="promoType" onchange="calcTotalOrder()">
                  <option value="Ninguno">Sin descuento</option>
                  <option value="Monto Fijo">Descuento en Monto ($)</option>
                  <option value="Porcentaje">Descuento Porcentaje (%)</option>
                  <option value="Regalo Promo">🎁 Regalo Promoción (100% Gratis)</option>
                  <option value="Canje Puntos">⭐ Canje de Puntos ($10 c/u)</option>
                </select>
              </div>
              <div class="form-group" style="margin-bottom:0;">
                <label>Valor Descuento ($ o %)</label>
                <input type="number" id="discountValue" value="0" min="0" oninput="calcTotalOrder()">
              </div>
              <div class="form-group" style="margin-bottom:0;">
                <label>Monto Final a Cobrar ($)</label>
                <input type="text" id="finalTotalDisplay" style="font-weight: 900; font-size: 1.3rem; color: var(--green-cash); background: #000;" readonly value="$10.000">
              </div>
            </div>
            <div style="margin-top: 8px; font-size: 0.78rem; color: var(--chalk-muted);" id="pointsEarnedPreview">
              ✨ Esta compra otorgará <strong>100 puntos de fidelidad</strong> al cliente.
            </div>
          </div>

          <button type="submit" class="btn btn-orange" style="width: 100%; font-size: 1.05rem; padding: 14px;">
            ✅ REGISTRAR VENTA, ASIGNAR PUNTOS Y RECALCULAR STOCK
          </button>
        </form>
      </div>
    </div>

    <!-- PANEL 2: CLIENTES & PUNTOS -->
    <div id="tab-clientes" class="tab-panel">
      <div class="card" style="border-left: 4px solid #25d366;">
        <div class="card-title">
          <span>📲 Enlace de Registro y Fidelización por WhatsApp</span>
          <span style="font-size: 0.85rem; color: #25d366;">Fideliza a tus clientes</span>
        </div>
        <p style="font-size: 0.85rem; color: var(--chalk-muted); margin-bottom: 14px;">
          Envía una invitación directa al WhatsApp del cliente para registrarlo en el Club Zetas con 50 puntos de bienvenida.
        </p>
        <div class="grid-3">
          <div class="form-group">
            <label>WhatsApp del Cliente (ej: 987654321)</label>
            <input type="tel" id="waClientPhone" placeholder="56987654321">
          </div>
          <div class="form-group">
            <label>Nombre del Cliente (Opcional)</label>
            <input type="text" id="waClientName" placeholder="Ej: Carlos">
          </div>
          <div class="form-group" style="display: flex; align-items: flex-end;">
            <button type="button" class="btn btn-green" style="width: 100%;" onclick="sendWhatsAppInvite()">
              🟢 Enviar Invitación WhatsApp
            </button>
          </div>
        </div>
      </div>

      <div class="grid-2">
        <div class="card">
          <div class="card-title">Inscribir / Editar Cliente</div>
          <form id="newClientForm" onsubmit="saveClientManual(event)">
            <div class="form-group">
              <label>Nombre Completo</label>
              <input type="text" id="regClientName" required placeholder="Ej: Ricardo Pereda">
            </div>
            <div class="form-group">
              <label>Teléfono / WhatsApp</label>
              <input type="tel" id="regClientPhone" required placeholder="+56 9 1234 5678">
            </div>
            <div class="form-group">
              <label>Dirección Habitual de Entrega</label>
              <input type="text" id="regClientAddress" required placeholder="Calle, Villa, N°, Referencia">
            </div>
            <div class="grid-2">
              <div class="form-group">
                <label>Cumpleaños (Opcional)</label>
                <input type="date" id="regClientBirthday">
              </div>
              <div class="form-group">
                <label>Puntos Iniciales</label>
                <input type="number" id="regClientPoints" value="50">
              </div>
            </div>
            <button type="submit" class="btn btn-orange" style="width:100%;">⭐ Guardar en Directorio</button>
          </form>
        </div>

        <div class="card">
          <div class="card-title">
            <span>Directorio de Clientes</span>
            <span style="font-size: 0.85rem; color: var(--gold-vip);" id="totalClientsCount">0 clientes</span>
          </div>
          <div class="table-container" style="max-height: 380px;">
            <table>
              <thead>
                <tr>
                  <th>Cliente</th>
                  <th>Teléfono</th>
                  <th>Dirección</th>
                  <th>Puntos</th>
                  <th>Acción</th>
                </tr>
              </thead>
              <tbody id="clientsTableBody"></tbody>
            </table>
          </div>
        </div>
      </div>
    </div>

    <!-- PANEL 3: GASTOS -->
    <div id="tab-gastos" class="tab-panel">
      <div class="card">
        <div class="card-title">Salida de Dinero / Gasto de Caja Chica</div>
        <form id="expenseForm" onsubmit="saveExpense(event)">
          <div class="grid-2">
            <div class="form-group">
              <label>Categoría de Gasto</label>
              <select id="expenseCategory" required>
                <option value="Insumos Frescos">Insumos Frescos (Verduras, Queso, Harina)</option>
                <option value="Bebidas y Reposición">Bebidas, Monster, Jugos (Reposición)</option>
                <option value="Packaging">Cajas de Pizza y Embalajes</option>
                <option value="Gas y Combustible">Gas / Bencina Repartidor</option>
                <option value="Pago Turno / Personal">Pago Repartidor / Ayudante Extra</option>
                <option value="Retiro Dueño">Retiro de Efectivo de Caja</option>
                <option value="Otros">Otros</option>
              </select>
            </div>
            <div class="form-group">
              <label>Monto Pagado ($)</label>
              <input type="number" id="expenseAmount" placeholder="Ej: 15000" min="100" required>
            </div>
          </div>
          <div class="grid-2">
            <div class="form-group">
              <label>Medio de Pago</label>
              <select id="expensePayment">
                <option value="Efectivo de Caja">Efectivo de Caja Chica</option>
                <option value="Transferencia">Transferencia Cuenta Negocio</option>
              </select>
            </div>
            <div class="form-group">
              <label>Descripción / Detalle</label>
              <input type="text" id="expenseDetail" placeholder="Ej: Reposición de Monster y cajas familiares" required>
            </div>
          </div>
          <button type="submit" class="btn btn-orange" style="width:100%;">💸 Registrar Egreso</button>
        </form>
      </div>
    </div>

    <!-- PANEL 4: HISTORIAL & EXPORTACIÓN EXCEL -->
    <div id="tab-registros" class="tab-panel">
      <div class="card">
        <div class="card-title">
          <span>Ventas Registradas en el Turno</span>
          <div style="display:flex; gap:8px;">
            <button class="btn btn-green" onclick="exportToExcelNative()">📊 Descargar Excel (.XLSX)</button>
            <button class="btn btn-ghost" onclick="copySalesForSheets()">📋 Copiar para Sheets</button>
          </div>
        </div>
        <div class="table-container">
          <table>
            <thead>
              <tr>
                <th>Hora</th>
                <th>Cliente</th>
                <th>Canal</th>
                <th>Pizza (3 Ingr.)</th>
                <th>Adicionales</th>
                <th>Bebidas/Otros</th>
                <th>Descuento / Promo</th>
                <th>Total</th>
                <th>Puntos</th>
                <th>Acción</th>
              </tr>
            </thead>
            <tbody id="salesTableBody"></tbody>
          </table>
        </div>
      </div>

      <div class="card">
        <div class="card-title">Gastos del Turno</div>
        <div class="table-container">
          <table>
            <thead>
              <tr>
                <th>Hora</th>
                <th>Categoría</th>
                <th>Detalle</th>
                <th>Medio</th>
                <th>Monto</th>
                <th>Acción</th>
              </tr>
            </thead>
            <tbody id="expensesTableBody"></tbody>
          </table>
        </div>
        <div style="margin-top: 16px;">
          <button class="btn btn-danger" onclick="clearCurrentShiftData()">⚠️ Reiniciar Todo el Turno</button>
        </div>
      </div>
    </div>

    <!-- PANEL 5: DASHBOARD DE DECISIONES -->
    <div id="tab-dashboard" class="tab-panel">
      <div class="kpi-grid">
        <div class="kpi-card" style="border-left-color: var(--green-cash);">
          <div class="kpi-title">Ingresos Totales (Ventas)</div>
          <div class="kpi-value" id="kpiTotalIncome">$0</div>
          <div class="kpi-sub" id="kpiOrdersCount">0 pedidos registrados</div>
        </div>
        <div class="kpi-card" style="border-left-color: var(--red-alert);">
          <div class="kpi-title">Gastos Caja Chica</div>
          <div class="kpi-value" id="kpiTotalExpenses">$0</div>
          <div class="kpi-sub" id="kpiExpensesCount">0 salidas de caja</div>
        </div>
        <div class="kpi-card" style="border-left-color: #38bdf8;">
          <div class="kpi-title">Margen Neto Operativo</div>
          <div class="kpi-value" id="kpiNetMargin">$0</div>
          <div class="kpi-sub" id="kpiMarginPercent">0% margen operacional</div>
        </div>
        <div class="kpi-card" style="border-left-color: var(--gold-vip);">
          <div class="kpi-title">Ticket Promedio</div>
          <div class="kpi-value" id="kpiAvgTicket">$0</div>
          <div class="kpi-sub">Por cliente atendido</div>
        </div>
        <div class="kpi-card" style="border-left-color: #ec4899;">
          <div class="kpi-title">Puntos Entregados</div>
          <div class="kpi-value" id="kpiPointsAwarded">0 pts</div>
          <div class="kpi-sub">Fidelización activa</div>
        </div>
      </div>

      <div class="grid-2">
        <div class="card">
          <div class="card-title">Rotación de Insumos Base (Arma tu Pizza)</div>
          <p style="font-size: 0.78rem; color: var(--chalk-muted); margin-bottom: 12px;">Ingredientes más demandados para compras mayoristas y evitar mermas.</p>
          <div id="baseIngredientsRanking"></div>
        </div>
        <div>
          <div class="card">
            <div class="card-title">Ventas de Bebidas, Monster & Otros</div>
            <div id="beveragesRanking"></div>
          </div>
          <div class="card">
            <div class="card-title">Adicionales Premium (+$2.000)</div>
            <div id="extraIngredientsRanking"></div>
          </div>
        </div>
      </div>
    </div>
  </div>

  <script>
    const BASE_INGREDIENTS = [
      "Pollo", "Carne", "Jamón", "Chorizo", "Tomate", "Champiñón",
      "Choclo", "Espárrago", "Toque de Pesto", "Palmito", "Cebolla Morada",
      "Albahaca", "Cebolla Caramelizada", "Rúcula", "Pimentón", "Aceituna",
      "Salame", "Pepperoni"
    ];

    const EXTRA_INGREDIENTS = [
      "Jamón Serrano", "Tomate Cherry", "Queso Parmesano",
      "Tocino", "Extra Queso", "Camarón Salteado"
    ];

    const BEVERAGES_CATALOG = [
      { id: "coca_15", name: "Coca-Cola 1.5L", price: 2500 },
      { id: "coca_lata", name: "Coca / Fanta / Sprite Lata", price: 1500 },
      { id: "monster", name: "Energética Monster", price: 2500 },
      { id: "redbull", name: "Red Bull", price: 2200 },
      { id: "jugo_natural", name: "Jugo Néctar 1.5L", price: 2000 },
      { id: "agua_mineral", name: "Agua Mineral 500cc", price: 1200 },
      { id: "papas_fritas", name: "Porción Papas Fritas", price: 3500 },
      { id: "salsa_ajo", name: "Salsa de Ajo Casera", price: 1000 }
    ];

    const PIZZA_BASE_PRICE = 9000;
    const EXTRA_PRICE = 2000;

    let salesData = JSON.parse(localStorage.getItem("zetas_sales_v3")) || [];
    let expensesData = JSON.parse(localStorage.getItem("zetas_expenses_v3")) || [];
    let clientsData = JSON.parse(localStorage.getItem("zetas_clients_v3")) || [
      { name: "Juan Pérez", phone: "+56991122334", address: "Av. Bisquertt 450, Rengo", points: 250, birthday: "1992-05-14" },
      { name: "Carolina Morales", phone: "+56988776655", address: "Villa San Desiderio Pje 3 #12", points: 410, birthday: "1988-11-20" }
    ];

    window.addEventListener("DOMContentLoaded", () => {
      renderBaseIngredients();
      renderExtraIngredients();
      renderBeveragesSelector();
      updateAllViews();
    });

    function switchTab(tabId) {
      document.querySelectorAll(".tab-panel").forEach(p => p.classList.remove("active"));
      document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
      document.getElementById("tab-" + tabId).classList.add("active");
      event.currentTarget.classList.add("active");
      if (tabId === 'dashboard') renderDashboard();
    }

    function renderBaseIngredients() {
      const container = document.getElementById("baseIngredientsContainer");
      container.innerHTML = BASE_INGREDIENTS.map((ing, idx) => `
        <div class="choice-chip">
          <input type="checkbox" id="base_${idx}" value="${ing}" onchange="handleBaseChange()">
          <label class="chip-label" for="base_${idx}">${ing}</label>
        </div>
      `).join("");
    }

    function renderExtraIngredients() {
      const container = document.getElementById("extraIngredientsContainer");
      container.innerHTML = EXTRA_INGREDIENTS.map((ext, idx) => `
        <div class="choice-chip">
          <input type="checkbox" id="extra_${idx}" value="${ext}" onchange="handleExtraChange()">
          <label class="chip-label" for="extra_${idx}">+ ${ext}</label>
        </div>
      `).join("");
    }

    function renderBeveragesSelector() {
      const container = document.getElementById("beveragesContainer");
      container.innerHTML = BEVERAGES_CATALOG.map(prod => `
        <div class="products-item">
          <div>
            <div style="font-weight:700; font-size:0.8rem;">${prod.name}</div>
            <div style="font-size:0.75rem; color:var(--chalk-muted);">$${prod.price.toLocaleString('es-CL')}</div>
          </div>
          <div class="qty-controls">
            <button type="button" class="qty-btn" onclick="changeBevQty('${prod.id}', -1)">-</button>
            <span id="qty_${prod.id}" style="font-weight:800; min-width:18px; text-align:center;">0</span>
            <button type="button" class="qty-btn" onclick="changeBevQty('${prod.id}', 1)">+</button>
          </div>
        </div>
      `).join("");
    }

    const beverageQuantities = {};
    function changeBevQty(prodId, delta) {
      beverageQuantities[prodId] = Math.max(0, (beverageQuantities[prodId] || 0) + delta);
      document.getElementById(`qty_${prodId}`).innerText = beverageQuantities[prodId];
      calcTotalOrder();
    }

    function handleBaseChange() {
      const maxBaseIngredients = 5;

      const selected = Array.from(
        document.querySelectorAll('#baseIngredientsContainer input[type="checkbox"]:checked')
      );

      document.getElementById("baseCounter").innerText =
        `${selected.length} / ${maxBaseIngredients} seleccionados`;

      const allBaseInputs = document.querySelectorAll(
        '#baseIngredientsContainer input[type="checkbox"]'
      );

      if (selected.length >= maxBaseIngredients) {
        allBaseInputs.forEach(input => {
          if (!input.checked) {
            input.disabled = true;
            input.parentElement.classList.add("disabled");
          }
        });
      } else {
        allBaseInputs.forEach(input => {
          input.disabled = false;
          input.parentElement.classList.remove("disabled");
        });
      }

      calcTotalOrder();
    }

    function handleExtraChange() {
      const selected = document.querySelectorAll('#extraIngredientsContainer input[type="checkbox"]:checked');
      const totalExtras = selected.length;

      document.getElementById("extraCounter").innerText =
        `${totalExtras} adicionales (+$${(totalExtras * EXTRA_PRICE).toLocaleString('es-CL')})`;

      calcTotalOrder();
    }

    function calcTotalOrder() {
      const baseSelectedCount = document.querySelectorAll('#baseIngredientsContainer input[type="checkbox"]:checked').length;
      const pizzaPrice = baseSelectedCount > 0 ? PIZZA_BASE_PRICE : 0;

      const extrasCount = document.querySelectorAll('#extraIngredientsContainer input[type="checkbox"]:checked').length;
      const extrasTotal = extrasCount * EXTRA_PRICE;

      let beveragesTotal = 0;
      BEVERAGES_CATALOG.forEach(p => {
        const qty = beverageQuantities[p.id] || 0;
        beveragesTotal += qty * p.price;
      });

      const deliveryFee = parseInt(document.getElementById("deliveryFee").value) || 0;
      const grossTotal = pizzaPrice + extrasTotal + beveragesTotal + deliveryFee;

      const promoType = document.getElementById("promoType").value;
      const discountInput = parseFloat(document.getElementById("discountValue").value) || 0;
      let discountAmount = 0;

      if (promoType === "Monto Fijo") {
        discountAmount = discountInput;
      } else if (promoType === "Porcentaje") {
        discountAmount = Math.round(grossTotal * (discountInput / 100));
      } else if (promoType === "Regalo Promo") {
        discountAmount = grossTotal;
      } else if (promoType === "Canje Puntos") {
        discountAmount = Math.min(grossTotal, discountInput * 10);
      }

      const netTotal = Math.max(0, grossTotal - discountAmount);
      document.getElementById("finalTotalDisplay").value = `$${netTotal.toLocaleString('es-CL')}`;

      const pointsEarned = Math.floor(netTotal / 100);
      document.getElementById("pointsEarnedPreview").innerHTML = `✨ Esta compra otorgará <strong>${pointsEarned} puntos de fidelidad</strong> al cliente.`;

      return { grossTotal, discountAmount, netTotal, pointsEarned };
    }

    function checkClientMatch() {
      const val = document.getElementById("clientName").value.toLowerCase().trim();
      const match = clientsData.find(c => c.name.toLowerCase().includes(val));
      const notice = document.getElementById("clientPointsNotice");
      if (match && val.length > 2) {
        document.getElementById("clientPhone").value = match.phone;
        document.getElementById("clientAddress").value = match.address;
        document.getElementById("clientPointsCount").innerText = match.points;
        notice.style.display = "block";
      } else {
        notice.style.display = "none";
      }
    }

    function openQuickClientSearch() {
      const search = prompt("Ingresa el nombre o teléfono del cliente a buscar:");
      if (!search) return;
      const found = clientsData.find(c => c.name.toLowerCase().includes(search.toLowerCase()) || c.phone.includes(search));
      if (found) {
        document.getElementById("clientName").value = found.name;
        document.getElementById("clientPhone").value = found.phone;
        document.getElementById("clientAddress").value = found.address;
        document.getElementById("clientPointsCount").innerText = found.points;
        document.getElementById("clientPointsNotice").style.display = "block";
      } else {
        alert("Cliente no encontrado en el Club. Puedes registrar sus datos en el pedido para agregarlo automáticamente.");
      }
    }

    function saveOrder(e) {
      e.preventDefault();
      const baseSelected = Array.from(document.querySelectorAll('#baseIngredientsContainer input[type="checkbox"]:checked')).map(i => i.value);
      const extraSelected = Array.from(document.querySelectorAll('#extraIngredientsContainer input[type="checkbox"]:checked')).map(i => i.value);

      const itemsBeverages = [];
      BEVERAGES_CATALOG.forEach(b => {
        const q = beverageQuantities[b.id] || 0;
        if (q > 0) itemsBeverages.push(`${b.name} (x${q})`);
      });

      if (baseSelected.length === 0 && itemsBeverages.length === 0) {
        alert("Debes seleccionar al menos una pizza o una bebida/extra para registrar el pedido.");
        return;
      }

      const clientName = document.getElementById("clientName").value.trim();
      const clientPhone = document.getElementById("clientPhone").value.trim();
      const clientAddress = document.getElementById("clientAddress").value.trim();

      const calc = calcTotalOrder();

      const newOrder = {
        id: Date.now(),
        date: new Date().toISOString().split('T')[0],
        time: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
        clientName: clientName,
        clientPhone: clientPhone,
        clientAddress: clientAddress,
        turno: document.getElementById("saleTurno").value,
        channel: document.getElementById("saleChannel").value,
        payment: document.getElementById("salePayment").value,
        baseIngredients: baseSelected,
        extraIngredients: extraSelected,
        beverages: itemsBeverages,
        promoType: document.getElementById("promoType").value,
        discountAmount: calc.discountAmount,
        total: calc.netTotal,
        pointsEarned: calc.pointsEarned
      };

      salesData.unshift(newOrder);
      localStorage.setItem("zetas_sales_v3", JSON.stringify(salesData));

      let existingClient = clientsData.find(c => c.phone === clientPhone || c.name.toLowerCase() === clientName.toLowerCase());
      if (existingClient) {
        existingClient.points += calc.pointsEarned;
        existingClient.address = clientAddress;
      } else {
        clientsData.push({
          name: clientName,
          phone: clientPhone,
          address: clientAddress,
          points: calc.pointsEarned + 50,
          birthday: ""
        });
      }
      localStorage.setItem("zetas_clients_v3", JSON.stringify(clientsData));

      document.getElementById("orderForm").reset();
      Object.keys(beverageQuantities).forEach(k => beverageQuantities[k] = 0);
      BEVERAGES_CATALOG.forEach(p => document.getElementById(`qty_${p.id}`).innerText = "0");
      handleBaseChange();
      handleExtraChange();
      document.getElementById("clientPointsNotice").style.display = "none";

      updateAllViews();
      alert(`✅ ¡Pedido registrado con éxito! Se han asignado ${calc.pointsEarned} puntos a ${clientName}.`);
    }

    function saveClientManual(e) {
      e.preventDefault();
      const name = document.getElementById("regClientName").value.trim();
      const phone = document.getElementById("regClientPhone").value.trim();
      const address = document.getElementById("regClientAddress").value.trim();
      const points = parseInt(document.getElementById("regClientPoints").value) || 0;
      const bday = document.getElementById("regClientBirthday").value;

      const exists = clientsData.find(c => c.phone === phone);
      if (exists) {
        exists.name = name;
        exists.address = address;
        exists.points += points;
        exists.birthday = bday;
      } else {
        clientsData.push({ name, phone, address, points, birthday: bday });
      }

      localStorage.setItem("zetas_clients_v3", JSON.stringify(clientsData));
      document.getElementById("newClientForm").reset();
      updateAllViews();
      alert("⭐ Cliente guardado en el Club de Fidelización.");
    }

    function sendWhatsAppInvite() {
      let phone = document.getElementById("waClientPhone").value.replace(/\D/g, "");
      const name = document.getElementById("waClientName").value.trim();
      if (!phone) {
        alert("Por favor ingresa el número telefónico del cliente.");
        return;
      }
      if (!phone.startsWith("56") && phone.length === 9) phone = "56" + phone;

      const greeting = name ? `¡Hola ${name}!` : "¡Hola!";
      const msg = `${greeting} Te invitamos a unirte al Club de Clientes de *Zetas Pizzería* 🍕✨.\n\nRegístrate para acumular puntos en cada pedido, canjear pizzas gratis, bebidas y promociones exclusivas en Rengo.\n\n👉 *Respóndenos con tus datos:*\n- Nombre Completo:\n- Dirección de despacho:\n- Fecha de cumpleaños:\n\n¡Al responder ya ganas 50 puntos de bienvenida! 🎉`;
      const url = `https://wa.me/${phone}?text=${encodeURIComponent(msg)}`;
      window.open(url, "_blank");
    }

    function saveExpense(e) {
      e.preventDefault();
      const amount = parseInt(document.getElementById("expenseAmount").value);
      if (!amount) return;

      const newExpense = {
        id: Date.now(),
        date: new Date().toISOString().split('T')[0],
        time: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
        category: document.getElementById("expenseCategory").value,
        amount: amount,
        payment: document.getElementById("expensePayment").value,
        detail: document.getElementById("expenseDetail").value
      };

      expensesData.unshift(newExpense);
      localStorage.setItem("zetas_expenses_v3", JSON.stringify(expensesData));
      document.getElementById("expenseForm").reset();
      updateAllViews();
      alert("💸 Gasto de caja registrado.");
    }

    function deleteSale(id) {
      if (confirm("¿Estás seguro de eliminar este pedido? Se descontará de las métricas e insumos.")) {
        salesData = salesData.filter(s => s.id !== id);
        localStorage.setItem("zetas_sales_v3", JSON.stringify(salesData));
        updateAllViews();
      }
    }

    function deleteExpense(id) {
      if (confirm("¿Deseas eliminar este registro de gasto?")) {
        expensesData = expensesData.filter(e => e.id !== id);
        localStorage.setItem("zetas_expenses_v3", JSON.stringify(expensesData));
        updateAllViews();
      }
    }

    function clearCurrentShiftData() {
      if (confirm("⚠️ ¿REINICIAR EL TURNO? Se limpiarán las ventas y gastos. Asegúrate de haber descargado el archivo Excel.")) {
        salesData = [];
        expensesData = [];
        localStorage.removeItem("zetas_sales_v3");
        localStorage.removeItem("zetas_expenses_v3");
        updateAllViews();
      }
    }

    function exportToExcelNative() {
      if (salesData.length === 0 && expensesData.length === 0) {
        alert("No hay ventas ni gastos registrados para exportar.");
        return;
      }

      const wb = XLSX.utils.book_new();

      const formattedSales = salesData.map(s => ({
        Fecha: s.date,
        Hora: s.time,
        Turno: s.turno,
        Cliente: s.clientName,
        Telefono: s.clientPhone,
        Direccion: s.clientAddress,
        Canal: s.channel,
        Medio_Pago: s.payment,
        Pizza_Ingredientes: s.baseIngredients.join(", "),
        Adicionales_2000: s.extraIngredients.join(", "),
        Bebidas_y_Otros: s.beverages.join(", "),
        Tipo_Promocion: s.promoType,
        Descuento_Aplicado: s.discountAmount,
        Total_Cobrado: s.total,
        Puntos_Asignados: s.pointsEarned
      }));
      const wsSales = XLSX.utils.json_to_sheet(formattedSales);
      XLSX.utils.book_append_sheet(wb, wsSales, "Ventas_Pizzeria");

      const formattedExpenses = expensesData.map(e => ({
        Fecha: e.date,
        Hora: e.time,
        Categoria: e.category,
        Detalle: e.detail,
        Medio_Pago: e.payment,
        Monto_Gasto: e.amount
      }));
      const wsExpenses = XLSX.utils.json_to_sheet(formattedExpenses);
      XLSX.utils.book_append_sheet(wb, wsExpenses, "Gastos_Caja");

      const wsClients = XLSX.utils.json_to_sheet(clientsData);
      XLSX.utils.book_append_sheet(wb, wsClients, "Club_Fidelizacion");

      const fileName = `Zetas_Pizzeria_${new Date().toISOString().split('T')[0]}.xlsx`;
      XLSX.writeFile(wb, fileName);
    }

    function copySalesForSheets() {
      if (salesData.length === 0) { alert("No hay ventas para copiar."); return; }
      let tsv = "Fecha\tHora\tCliente\tTelefono\tDireccion\tCanal\tPizza_Ingredientes\tAdicionales\tBebidas\tDescuento\tTotal\tPuntos\n";
      salesData.forEach(s => {
        tsv += `${s.date}\t${s.time}\t${s.clientName}\t${s.clientPhone}\t${s.clientAddress}\t${s.channel}\t${s.baseIngredients.join(", ")}\t${s.extraIngredients.join(", ")}\t${s.beverages.join(", ")}\t${s.discountAmount}\t${s.total}\t${s.pointsEarned}\n`;
      });
      navigator.clipboard.writeText(tsv).then(() => {
        alert("📋 ¡Copiado! Pega directamente con Ctrl + V en tu Google Sheets.");
      });
    }

    function updateAllViews() {
      const sTbody = document.getElementById("salesTableBody");
      sTbody.innerHTML = salesData.length === 0 ? `<tr><td colspan="10" style="text-align:center; color:var(--chalk-muted);">No hay ventas registradas aún.</td></tr>` :
        salesData.map(s => `
          <tr>
            <td>${s.time}</td>
            <td><strong>${s.clientName}</strong><br><span style="font-size:0.75rem; color:var(--chalk-muted);">${s.clientAddress}</span></td>
            <td>${s.channel}</td>
            <td>${s.baseIngredients.join(" + ") || "Sin pizza"}</td>
            <td>${s.extraIngredients.length > 0 ? `<span style="color:#ffb74d;">${s.extraIngredients.join(", ")}</span>` : '-'}</td>
            <td>${s.beverages.join("<br>") || '-'}</td>
            <td>${s.discountAmount > 0 ? `<span style="color:var(--red-alert);">-$${s.discountAmount.toLocaleString('es-CL')} (${s.promoType})</span>` : 'Sin dcto.'}</td>
            <td style="font-weight:800; color:var(--green-cash);">$${s.total.toLocaleString('es-CL')}</td>
            <td><span class="tag tag-points">+${s.pointsEarned} pts</span></td>
            <td><button class="btn btn-danger" onclick="deleteSale(${s.id})" style="padding:4px 8px; font-size:0.75rem;">🗑️</button></td>
          </tr>
        `).join("");

      const eTbody = document.getElementById("expensesTableBody");
      eTbody.innerHTML = expensesData.length === 0 ? `<tr><td colspan="6" style="text-align:center; color:var(--chalk-muted);">No hay gastos registrados.</td></tr>` :
        expensesData.map(e => `
          <tr>
            <td>${e.time}</td>
            <td><strong>${e.category}</strong></td>
            <td>${e.detail}</td>
            <td>${e.payment}</td>
            <td style="font-weight:800; color:var(--red-alert);">$${e.amount.toLocaleString('es-CL')}</td>
            <td><button class="btn btn-danger" onclick="deleteExpense(${e.id})" style="padding:4px 8px; font-size:0.75rem;">🗑️</button></td>
          </tr>
        `).join("");

      document.getElementById("totalClientsCount").innerText = `${clientsData.length} inscritos`;
      const cTbody = document.getElementById("clientsTableBody");
      cTbody.innerHTML = clientsData.sort((a,b) => b.points - a.points).map(c => `
        <tr>
          <td><strong>${c.name}</strong></td>
          <td>${c.phone}</td>
          <td>${c.address}</td>
          <td><span class="tag tag-points">⭐ ${c.points} pts</span></td>
          <td><button class="btn btn-ghost" style="padding:3px 6px; font-size:0.75rem;" onclick="applyClientOrder('${c.phone}')">Seleccionar</button></td>
        </tr>
      `).join("");

      renderDashboard();
    }

    function applyClientOrder(phone) {
      const c = clientsData.find(cli => cli.phone === phone);
      if (c) {
        document.getElementById("clientName").value = c.name;
        document.getElementById("clientPhone").value = c.phone;
        document.getElementById("clientAddress").value = c.address;
        document.getElementById("clientPointsCount").innerText = c.points;
        document.getElementById("clientPointsNotice").style.display = "block";
        switchTab('ventas');
      }
    }

    function renderDashboard() {
      const totalSales = salesData.reduce((acc, curr) => acc + curr.total, 0);
      const totalExpenses = expensesData.reduce((acc, curr) => acc + curr.amount, 0);
      const netMargin = totalSales - totalExpenses;
      const marginPct = totalSales > 0 ? ((netMargin / totalSales) * 100).toFixed(1) : 0;
      const avgTicket = salesData.length > 0 ? Math.round(totalSales / salesData.length) : 0;
      const totalPoints = salesData.reduce((acc, curr) => acc + curr.pointsEarned, 0);

      document.getElementById("kpiTotalIncome").innerText = `$${totalSales.toLocaleString('es-CL')}`;
      document.getElementById("kpiOrdersCount").innerText = `${salesData.length} pedidos realizados`;
      document.getElementById("kpiTotalExpenses").innerText = `$${totalExpenses.toLocaleString('es-CL')}`;
      document.getElementById("kpiExpensesCount").innerText = `${expensesData.length} salidas de caja`;
      document.getElementById("kpiNetMargin").innerText = `$${netMargin.toLocaleString('es-CL')}`;
      document.getElementById("kpiMarginPercent").innerText = `${marginPct}% margen operacional`;
      document.getElementById("kpiAvgTicket").innerText = `$${avgTicket.toLocaleString('es-CL')}`;
      document.getElementById("kpiPointsAwarded").innerText = `${totalPoints} pts`;

      const baseCounts = {};
      BASE_INGREDIENTS.forEach(ing => baseCounts[ing] = 0);
      salesData.forEach(s => s.baseIngredients && s.baseIngredients.forEach(i => baseCounts[i] = (baseCounts[i] || 0) + 1));
      const sortedBase = Object.entries(baseCounts).sort((a,b) => b[1] - a[1]);
      const maxBase = Math.max(...Object.values(baseCounts), 1);

      document.getElementById("baseIngredientsRanking").innerHTML = sortedBase.map(([ing, cnt]) => `
        <div class="bar-row">
          <div class="bar-header"><span>${ing}</span><strong style="color:var(--orange-accent);">${cnt} veces</strong></div>
          <div class="bar-track"><div class="bar-fill" style="width:${(cnt / maxBase) * 100}%;"></div></div>
        </div>
      `).join("");

      const bevCounts = {};
      salesData.forEach(s => {
        if (s.beverages) {
          s.beverages.forEach(b => {
            const name = b.split(" (x")[0];
            const qty = parseInt(b.split(" (x")[1]) || 1;
            bevCounts[name] = (bevCounts[name] || 0) + qty;
          });
        }
      });
      const sortedBev = Object.entries(bevCounts).sort((a,b) => b[1] - a[1]);
      document.getElementById("beveragesRanking").innerHTML = sortedBev.length === 0 ? `<p style="font-size:0.8rem; color:var(--chalk-muted);">No hay bebidas vendidas aún.</p>` :
        sortedBev.map(([bName, cnt]) => `
          <div style="display:flex; justify-content:space-between; padding:6px 0; border-bottom:1px solid var(--border-chalk); font-size:0.85rem;">
            <span>🥤 ${bName}</span>
            <strong style="color:var(--green-cash);">${cnt} un.</strong>
          </div>
        `).join("");

      const extraCounts = {};
      EXTRA_INGREDIENTS.forEach(ext => extraCounts[ext] = 0);
      salesData.forEach(s => s.extraIngredients && s.extraIngredients.forEach(e => extraCounts[e] = (extraCounts[e] || 0) + 1));
      const sortedExtra = Object.entries(extraCounts).sort((a,b) => b[1] - a[1]);
      document.getElementById("extraIngredientsRanking").innerHTML = sortedExtra.map(([ext, cnt]) => `
        <div style="display:flex; justify-content:space-between; padding:6px 0; border-bottom:1px solid var(--border-chalk); font-size:0.85rem;">
          <span>🥓 ${ext}</span>
          <strong style="color:#ffb74d;">${cnt} ($${(cnt * EXTRA_PRICE).toLocaleString('es-CL')})</strong>
        </div>
      `).join("");
    }
  </script>
</body>
</html>
