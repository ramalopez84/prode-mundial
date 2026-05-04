import React, { useState, useEffect, useCallback } from "react";

// Nueva URL de implementación
const API_URL = "https://script.google.com/macros/s/AKfycbxR1iWTXd8QqU3eAyud07xkYUp6z3lW781djgiAddqeBGqo0VHwN9mIMcgSj7g8darHqg/exec";

export default function MundialApp() {
  const [vista, setVista] = useState("ingreso");
  const [datos, setDatos] = useState([]);
  const [cargando, setCargando] = useState(false);
  const [error, setError] = useState(null);

  // --- ESTILOS ---
  const stickerStyle = {
    filter: 'drop-shadow(3px 0 0 white) drop-shadow(-3px 0 0 white) drop-shadow(0 3px 0 white) drop-shadow(0 -3px 0 white) drop-shadow(5px 5px 12px rgba(0,0,0,0.4))',
    maxWidth: '220px',
    height: 'auto',
    marginBottom: '30px',
    transition: 'transform 0.3s ease'
  };

  const btnStyle = {
    padding: '12px 24px',
    margin: '10px',
    borderRadius: '8px',
    border: 'none',
    cursor: 'pointer',
    fontWeight: 'bold',
    backgroundColor: '#007bff',
    color: 'white'
  };

  // --- FUNCIONES DE BASE DE DATOS ---
  
  const cargarBase = useCallback(async (hoja = "Usuarios") => {
    setCargando(true);
    try {
      const res = await fetch(`${API_URL}?hoja=${hoja}`);
      const json = await res.json();
      setDatos(json.datos || []);
    } catch (e) {
      setError("Error al conectar con la base de datos");
    } finally {
      setCargando(false);
    }
  }, []);

  const ejecutarAccionAdmin = async (payload) => {
    setCargando(true);
    try {
      // Usamos no-cors para evitar bloqueos de Google al hacer POST
      await fetch(API_URL, {
        method: "POST",
        mode: "no-cors",
        body: JSON.stringify(payload)
      });
      
      alert("Operación realizada. Los cambios se verán reflejados en breve.");
      // Recargamos los datos para confirmar
      setTimeout(() => cargarBase(payload.hoja), 1500);
    } catch (e) {
      alert("Hubo un problema al ejecutar la acción.");
    } finally {
      setCargando(false);
    }
  };

  const descargarExcel = () => {
    if (datos.length === 0) return alert("No hay datos para exportar.");
    
    // Generamos un CSV (compatible con Excel)
    const csvRows = datos.map(fila => fila.join(","));
    const csvContent = "data:text/csv;charset=utf-8,\uFEFF" + csvRows.join("\n");
    const encodedUri = encodeURI(csvContent);
    const link = document.createElement("a");
    link.setAttribute("href", encodedUri);
    link.setAttribute("download", "Prode_Mundial_Data.csv");
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  return (
    <div style={{ textAlign: 'center', padding: '20px', fontFamily: 'Segoe UI, Tahoma, Geneva, Verdana, sans-serif', color: '#333' }}>
      
      {/* VISTA: INGRESO */}
      {vista === "ingreso" && (
        <div style={{ marginTop: '50px' }}>
          <img 
            src="https://upload.wikimedia.org/wikipedia/commons/thumb/1/10/AFA_logo.svg/1200px-AFA_logo.svg.png" 
            alt="Logo Mundial" 
            style={stickerStyle} 
          />
          <h1 style={{ fontSize: '2.5rem', color: '#1a3a5a' }}>Prode Copa del Mundo</h1>
          <div style={{ marginTop: '30px' }}>
            <button style={btnStyle} onClick={() => setVista("pronosticos")}>
              📝 Cargar Pronósticos
            </button>
            <button 
              style={{ ...btnStyle, backgroundColor: '#6c757d' }} 
              onClick={() => { setVista("admin"); cargarBase(); }}
            >
              ⚙️ Administrador
            </button>
          </div>
        </div>
      )}

      {/* VISTA: PRONÓSTICOS */}
      {vista === "pronosticos" && (
        <div>
          <button style={{ ...btnStyle, backgroundColor: '#f8f9fa', color: '#333' }} onClick={() => setVista("ingreso")}>
            ← Volver
          </button>
          <h2>Mis Predicciones</h2>
          <div style={{ border: '1px solid #ddd', padding: '20px', borderRadius: '10px', display: 'inline-block' }}>
            <p><strong>Partido 1:</strong> Argentina vs Arabia Saudita</p>
            <input type="number" placeholder="0" style={{ width: '40px', textAlign: 'center' }} /> - 
            <input type="number" placeholder="0" style={{ width: '40px', textAlign: 'center' }} />
            <br /><br />
            <button style={{ ...btnStyle, backgroundColor: '#28a745' }}>Guardar Resultados</button>
          </div>
        </div>
      )}

      {/* VISTA: ADMINISTRADOR */}
      {vista === "admin" && (
        <div>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px' }}>
            <button style={{ ...btnStyle, backgroundColor: '#f8f9fa', color: '#333' }} onClick={() => setVista("ingreso")}>
              ← Volver
            </button>
            <h2>Panel de Control</h2>
            <button onClick={descargarExcel} style={{ ...btnStyle, backgroundColor: '#1d6f42' }}>
              📊 Descargar Excel
            </button>
          </div>

          <div style={{ backgroundColor: '#f1f3f5', padding: '15px', borderRadius: '10px', marginBottom: '20px' }}>
            <h4>Acciones Rápidas</h4>
            <button 
              onClick={() => ejecutarAccionAdmin({ accion: "fusionar", idViejo: "Bachi", idNuevo: "Mono", hoja: "Usuarios" })}
              style={{ ...btnStyle, backgroundColor: '#ffc107', color: '#000' }}
            >
              🔀 Fusionar Bachi → Mono
            </button>
          </div>

          {cargando ? <p>Procesando datos...</p> : (
            <div style={{ overflowX: 'auto' }}>
              <table border="1" style={{ width: '100%', borderCollapse: 'collapse', marginTop: '10px' }}>
                <thead style={{ backgroundColor: '#1a3a5a', color: 'white' }}>
                  <tr>
                    {datos[0]?.map((h, i) => <th key={i} style={{ padding: '10px' }}>{h}</th>)}
                    <th style={{ padding: '10px' }}>Gestión</th>
                  </tr>
                </thead>
                <tbody>
                  {datos.slice(1).map((fila, i) => (
                    <tr key={i} style={{ backgroundColor: i % 2 === 0 ? '#fff' : '#f9f9f9' }}>
                      {fila.map((c, j) => <td key={j} style={{ padding: '8px' }}>{c}</td>)}
                      <td style={{ padding: '8px' }}>
                        <button 
                          onClick={() => ejecutarAccionAdmin({ accion: "eliminar", id: fila[0], hoja: "Usuarios" })}
                          style={{ backgroundColor: '#dc3545', color: 'white', border: 'none', padding: '5px 10px', borderRadius: '4px', cursor: 'pointer' }}
                        >
                          Eliminar
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      )}
    </div>
  );
}