#!/usr/bin/env python3
"""Genera manifest.xml y config.js con tus datos.

Uso:
    python configurar.py --url https://TUUSUARIO.github.io/adjuntos-hilo \
                         --client-id xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx \
                         --tenant-id yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy

--url        URL base HTTPS donde quedan alojados estos archivos (sin barra final).
--client-id  "Id. de aplicación (cliente)" del registro en Entra ID.
--tenant-id  "Id. de directorio (inquilino)" del registro en Entra ID.

Se puede ejecutar tantas veces como haga falta; el Id del complemento (GUID)
se conserva entre ejecuciones en el archivo .app-id para que Outlook lo
reconozca como el mismo complemento.
"""
import argparse
import pathlib
import re
import uuid

AQUI = pathlib.Path(__file__).resolve().parent


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--url", required=True, help="URL base HTTPS del complemento, sin barra final")
    p.add_argument("--client-id", required=True)
    p.add_argument("--tenant-id", required=True)
    args = p.parse_args()

    url = args.url.rstrip("/")
    if not url.startswith("https://"):
        p.error("--url debe empezar por https://")
    for nombre, valor in (("--client-id", args.client_id), ("--tenant-id", args.tenant_id)):
        if not re.fullmatch(r"[0-9a-fA-F-]{36}", valor):
            p.error(f"{nombre} no parece un GUID válido: {valor}")

    archivo_id = AQUI / ".app-id"
    if archivo_id.exists():
        app_id = archivo_id.read_text().strip()
    else:
        app_id = str(uuid.uuid4())
        archivo_id.write_text(app_id + "\n")

    sustituciones = {
        "__BASE_URL__": url,
        "__CLIENT_ID__": args.client_id,
        "__TENANT_ID__": args.tenant_id,
        "__APP_ID__": app_id,
    }

    for plantilla, salida in (("manifest.template.xml", "manifest.xml"), ("config.template.js", "config.js")):
        texto = (AQUI / plantilla).read_text(encoding="utf-8")
        for k, v in sustituciones.items():
            texto = texto.replace(k, v)
        (AQUI / salida).write_text(texto, encoding="utf-8")
        print(f"Generado {salida}")

    dominio = url.split("//", 1)[1].split("/", 1)[0]
    print()
    print("Comprueba en Entra ID > Autenticación > Aplicación de página única que tienes estos URI de redirección:")
    print(f"  brk-multihub://{dominio}")
    print(f"  {url}/taskpane.html")
    print()
    print("Sube toda la carpeta al hosting y luego carga manifest.xml en Outlook")
    print("(Mis complementos > Complementos personalizados > Agregar desde archivo).")


if __name__ == "__main__":
    main()
