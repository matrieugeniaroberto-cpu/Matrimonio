# Boda Yuyito & Roberto

App de planificación de matrimonio con dos eventos, invitados, presupuesto,
proveedores, catering, mesas, cronograma, scrapbook, tablero de inspiración
y guardado automático en tiempo real vía Supabase.

## 1. Configurar Supabase

1. Entra a tu proyecto en [supabase.com](https://supabase.com) (o crea uno nuevo, es gratis).
2. Ve a **SQL Editor → New query**, pega el contenido del archivo `supabase.sql` de esta carpeta y dale **Run**. Esto crea la tabla `wedding_data`, sus permisos, y activa Realtime.
3. Ve a **Project Settings → API**. Copia:
   - **Project URL** → esto es `VITE_SUPABASE_URL`
   - **anon public key** → esto es `VITE_SUPABASE_ANON_KEY`

## 2. Probar en tu computador (opcional pero recomendado)

```bash
npm install
cp .env.example .env
# edita .env y pega tu URL y anon key
npm run dev
```

Abre la URL que muestra la terminal (normalmente `http://localhost:5173`). Prueba
agregar un invitado y revisa en Supabase → **Table Editor → wedding_data** que el
campo `data` se haya actualizado.

## 3. Subir el código a GitHub

```bash
git init
git add .
git commit -m "App de boda"
git branch -M main
git remote add origin https://github.com/TU-USUARIO/TU-REPO.git
git push -u origin main
```

(Crea el repo vacío en GitHub primero, sin README, y reemplaza la URL de arriba
por la tuya.) El archivo `.gitignore` ya excluye `node_modules` y tu `.env`,
así que tu anon key no queda en el historial por accidente si prefieres
cargarla solo en Vercel.

## 4. Desplegar en Vercel

1. En [vercel.com](https://vercel.com) → **Add New → Project** → importa el repo de GitHub que acabas de crear.
2. Vercel detecta Vite automáticamente (Framework Preset: Vite). No cambies nada del build.
3. En **Environment Variables** agrega:
   - `VITE_SUPABASE_URL` = tu Project URL
   - `VITE_SUPABASE_ANON_KEY` = tu anon key
4. **Deploy**. En 1–2 minutos te da una URL tipo `boda-yuyito-roberto.vercel.app` — esa es la que abres en el celular y el computador, y la que le pasas a Roberto.

Cada vez que hagas `git push` a `main`, Vercel vuelve a desplegar solo.

## Notas

- Los datos se guardan automáticamente ~0.7s después de cualquier cambio, y se
  sincronizan en vivo entre dispositivos abiertos al mismo tiempo (Realtime).
- El botón "Descargar respaldo (JSON)" en Configuración sigue sirviendo como
  copia offline adicional.
- La URL de la app funciona sin login — trátala como privada de la pareja, no
  la publiques en redes.
