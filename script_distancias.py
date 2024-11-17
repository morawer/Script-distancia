import pandas as pd
import googlemaps
# Inicialización del cliente de Google Maps

API_KEY = ""  # Reemplaza con tu API Key de Google Maps
gmaps = googlemaps.Client(key=API_KEY)

# Cargar el archivo Excel
input_file = "info_suppliers.xlsx"
output_file = "info_suppliers_with_distances.xlsx"
data = pd.read_excel(input_file)

# Dirección base
origin_address = "Calle Montecarlo, 14, Fuenlabrada, Madrid"

# Combinar las columnas de dirección en una sola dirección completa


def combine_address(row):
    address_parts = [
        str(row.get(col, "")).strip() for col in [
            "Postal Address line 1",
            "Postal Address line 2",
            "Postal Address line 3",
            "Postal Address line 4"
        ]
    ]
    return ", ".join(filter(None, address_parts))


data["Full Address"] = data.apply(combine_address, axis=1)

# Función para calcular la distancia


def calculate_distance(destination):
    if not destination:
        return None
    try:
        result = gmaps.distance_matrix(
            origins=origin_address,
            destinations=destination,
            mode="driving"
        )

        # Extraer la distancia en kilómetros
        distance = result["rows"][0]["elements"][0].get(
            "distance", {}).get("value", 0)
        print(f"{destination}: {distance / 1000} kilometros")
        return round(distance / 1000, 2)  # Convertir de metros a kilómetros
    except Exception as e:
        print(f"Error con dirección {destination}: {e}")
        return None


# Calcular las distancias
data["distance_kms"] = data["Full Address"].apply(calculate_distance)

# Guardar el archivo actualizado
data.to_excel(output_file, index=False)
print(f"Archivo guardado como {output_file}")
