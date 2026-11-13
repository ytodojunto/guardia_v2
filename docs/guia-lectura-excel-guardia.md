# Cómo se lee el Excel de la guardia (MULTIPAR)

Guía para que cualquier herramienta (página de Proyección, dashboard, etc.) interprete el Excel de la guardia igual que lo hacen los operadores. Está armada con las reglas que fuimos definiendo en el proyecto.

> **Estado:** reglas confirmadas por Chino salvo donde dice *(a confirmar)*. No pude abrir el archivo "MAYO 2026 1.xlsx" (llegó vacío), así que la guía sale de lo ya acordado y no de inspeccionar ese libro puntual. Conviene contrastarla con el archivo real.

---

## 1. Estructura del libro

Hojas del libro "GUARDIA 2026":

| Hoja | Qué contiene |
|---|---|
| **GUARDIA** | Vista principal: turno en lista, en viaje, francos, enfermos, movimientos, observaciones, teléfonos de lanchas |
| **SUBIDAS** | Buques programados que suben (a Rosario y puertos del norte) |
| **BAJADAS** | Buques programados que bajan |
| **CAMPANA** | Cajón de movimientos: Campana, Zárate, Delta Dock **y también** movimientos cortos (San Nicolás a San Lorenzo, Rosario a San Nicolás, Villa Constitución a Rosario). El encabezado dice "Campana" aunque el buque no vaya allí |
| **GUARDIA REMISES** | Remises |
| **NOVEDADES** | Novedades |

**Fuente a futuro:** el reporte oficial de la guardia en PDF (sistema) probablemente reemplace al Excel como fuente.

---

## 2. Leer por anclas, no por posición

Las filas **no son fijas**: el turno puede tener hasta 25 prácticos, así que los bloques se corren. Hay que ubicar cada sección buscando su rótulo en la **columna A (coincidencia exacta)**:

- `FRANCOS`
- `ENFERMOS`
- `OBSERVACIONES`
- `EN VIAJE A ROSARIO`
- `EN VIAJE DE ROSARIO`
- `GUARDIA ROSARIO` (sin "EN" en el medio)
- `EN VIAJE A CAMPANA`
- `MOVIMIENTOS`
- `TELEFONOS DE LAS LANCHAS`
- `PRESIDENTE:`
- `COORDINADOR:`

Cada bloque va desde su ancla hasta la siguiente. Si un ancla no aparece, la lectura tiene que **avisar** (warning visible) en vez de devolver datos vacíos en silencio.

---

## 3. Celdas combinadas

El libro usa muchas celdas combinadas. El valor vive en la **celda superior izquierda** del rango combinado; el resto viene vacío. Al leer, hay que "rellenar hacia abajo/derecha" el valor de la combinada, y al escribir (openpyxl) siempre hacerlo en la celda superior izquierda.

---

## 4. Marcas de un solo carácter

Vienen del Excel viejo y se conservan. Van en una **columna dedicada al comienzo de cada tabla** (salvo casos heredados donde quedan pegadas al texto).

- **GUARDIA (en turno):** `*` junto al nombre = el operador **contactó** a ese práctico. **No** es un cambio de turno.
- **SUBIDAS / BAJADAS / CAMPANA:** `*` = se contactó a la agencia; `X` = buque **despachado**.

---

## 5. Colores (dependen de la sección)

El mismo color significa cosas distintas según dónde esté. Se lee el color de fondo de la celda.

**Celda Buque/Agencia** en GUARDIA (en viaje a/de Rosario, Campana/Movimientos) y en SUBIDAS/BAJADAS/CAMPANA. **No aplica** a EN TURNO, GUARDIA EN ROSARIO ni FRANCOS:

| Color | Hex | Significado |
|---|---|---|
| Gris | `#D9D9D9` | CUANDO DESPACHAN (despachado) |
| Naranja | `#FFC000` | NO ES NUESTRO (servicio de otra empresa) |
| Verde | `#75A67C` | NOS CUBREN (otra empresa cubre un buque nuestro) |
| Amarillo fuerte | `#FFFF00` | EN FRANCO |

**Otras secciones:**

- **Lista EN TURNO:** naranja = prácticos ya despachados.
- **En viaje (íconos de práctico):** ícono sombreado de color = práctico de otra empresa cubriendo un buque de MULTIPAR.
- **Hoja SUBIDAS (lectura específica):** gris = despachados; naranja = autorizados; amarillo fuerte = barcos que cubrimos a otras empresas.

*(a confirmar)* En SUBIDAS/BAJADAS/CAMPANA no hay columna de práctico asignado, así que todavía no hay regla automática confirmada para distinguir "NO ES NUESTRO" de "NOS CUBREN" ahí; falta que Chino pase los códigos de color exactos.

---

## 6. Regla del catálogo de prácticos

MULTIPAR tiene **26 prácticos propios** (catálogo maestro). Cualquier nombre que aparezca en la guardia y **no** esté en ese catálogo es un caso de cobertura:

- práctico de otra empresa cubriendo un buque de MULTIPAR, o
- práctico de MULTIPAR cubriendo un buque de otra empresa.

Externos que cubren barcos sueltos y **no entran a la cola del turno**: Falcon (no está en turno permanente), Luis Campetella y los freelance Brautighan y Deleesneider.

---

## 7. Cómo se cuenta un despacho

Un **despacho se cuenta por práctico**: un buque con dos prácticos son dos despachos.

**Dotación base** (parámetros ajustables, a contrastar con el Reglamento):

- Zona Común a San Pedro y todo al norte (Arroyo Seco, San Lorenzo, Rosario…): 2 prácticos, subida y bajada.
- Campana hasta 8,84 m de calado: 1 de subida y 1 de bajada. Desde 8,84 m (incluido): 2.
- Zárate / Delta Dock: 2.
- Terminal Guazú: 2 de subida; 1 de bajada si baja por Martín García.
- Bajada desde San Lorenzo: 2.
- Nueva Palmira: subidas con 2 hasta el puerto de arribo.
- Los movimientos de la hoja CAMPANA normalmente van con 1.

Práctica real: normalmente salen con 2 desde el inicio. Solo en bajada, si la guardia está apretada, puede salir 1 y sumarse el segundo después. En subida siempre salen los 2.

---

## 8. Orden y disponibilidad de los prácticos

- El llamado se hace **siempre por orden de pizarra**: n.º 1, luego n.º 2, etc. No se saltea a nadie por estar en descanso.
- Los que vienen bajando se reincorporan **al final de la lista**, por orden de desembarco.
- **Disponible desde:** hora de desembarco + presentación (Anexo I del Reglamento, punto 1.3.1) + 6 h de descanso.
- Un práctico que sale de la lista por cualquier motivo (franco, enfermo) queda afuera como mínimo 24 h y vuelve con el número con que salió.
- **Francos:** rotación 14x14 en dos grupos (A y B), cambio a las 06:00. No es lineal: hay permutas entre colegas, así que la fuente confiable es la guardia misma, no la planilla de rotación.

---

## 9. Buques en viaje

- En la fila de un buque en viaje, el práctico **n.º 1** es el que baja primero (punto intermedio, ej. San Nicolás o Escobar); el **n.º 2** llega hasta el final.
- En las bajadas, uno desembarca normalmente en Escobar (pocas veces en Brazo Largo, o en Campana con mal tiempo previsto) y el otro sigue a Zona Común.

---

## 10. Tiempos de referencia

**Velocidades:** subida ~10 nudos, bajada ~12 nudos.

| Trayecto | Tiempo |
|---|---|
| Zona Común a San Nicolás (subida) | 16 h |
| Zona Común a San Lorenzo (subida) | 22 h |
| Campana a San Lorenzo (subida) | 19 h |
| Campana a Zona Común | 6 h 30 min |
| Nueva Palmira a San Lorenzo | 16 h |
| San Lorenzo a Nueva Palmira | 14 h |
| Rosario a San Lorenzo (movimiento) | 2 h |
| Villa Constitución a Rosario | 2 h 30 min |

Las distancias salen de la tabla de distancias al final de la tarifa de pilotaje. Los tiempos de presentación en puerto salen de la tabla del Anexo I.

**Horas de inicio:** en subida, una vez que el buque confirma su hora, no se demora más de 1 o 2 h. En bajada se cumple la hora de pedido a bordo en la terminal y el inicio formal es ~1 h después en promedio, según la terminal.

---

## 11. Vocabulario

- **Práctico de río:** cualquier socio de MULTIPAR (Río Paraná). **Práctico de puerto:** el colega que hace las maniobras de atraque y desatraque.
- **Despacho:** salida de un práctico a un servicio.
- **Franco:** día libre. **Turno:** lista de prácticos de guardia.
- **GFH (grupo fecha hora):** código de Prefectura con día y hora de autorización para que un buque siga navegando a un puerto.
- **Estacionario:** muelle donde amarra y sale la lancha que embarca y desembarca prácticos.
- **Rada:** zona de fondeo.
- **ZC:** Zona Común. **ESC:** Escobar.

---

## 12. Reporte PDF del sistema (futura fuente)

- "Pedido para" = hora en que el práctico está a bordo; el inicio (I) arranca a esa hora.
- `D1` / `D2` = desembarca práctico 1 / práctico 2.
- Trae también los barcos previstos ("Sin despachar").
- Entradas raras (por ejemplo el buque Zara de septiembre) son pruebas de quienes mantienen el sistema.

---

## 13. Compensación entre empresas

Si MULTIPAR no cubre un barco y se lo da a otra empresa, esa empresa en algún momento le pasa uno a MULTIPAR (y a la inversa). Conviene llevar el saldo por empresa.

---

## 14. Trampas conocidas al leer

- **Arrastrar/autocompletar en Excel** duplica filas. Los operadores tienen que copiar y borrar, no arrastrar.
- Hay formatos heredados donde una marca (`*`, `X`) quedó pegada al nombre o al buque. Normalizar antes de comparar.
- El título "CAMPANA" no garantiza que el buque vaya a Campana (ver sección 1).
- Un nombre que no está en el catálogo no es un error: es cobertura (sección 6).

## NOVEDADES OPERATIVAS: avisos del momento (v2.6)

Columnas: A FECHA, B INFORMO, C NOVEDADES, **D PRIORIDAD** (Urgente o Normal) y **E VIGENTE HASTA** (fecha y hora, ej. 14/11/2026 18:00).
Si E tiene fecha y hora y todavía no venció, la página muestra la novedad en un cuadro bajo el encabezado (urgentes en rojo, primero). Al vencer desaparece sola de ahí y queda en la pestaña Novedades. Si E está vacía es una novedad de registro. Si E solo tiene fecha, vale hasta las 23:59.
