import requests
from datetime import datetime
import xlwings as xw
from openpyxl import load_workbook


# Funktion til at hente data fra DMR API
def fetch_vehicle_data(registration_number, api_token):
    """
    Henter køretøjsdata fra DMR API baseret på nummerpladen.
    """
    url = f"https://api.nrpla.de/evaluations/{registration_number}"
    headers = {"Authorization": f"Bearer {api_token}"}
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()  # Tjek for HTTP-fejl
        return response.json()["data"][0]  # Returnér første datasæt
    except requests.exceptions.RequestException as e:
        raise Exception(f"API request failed: {e}")
    except Exception as e:
        raise Exception(f"Error fetching vehicle data: {e}")


# Funktion til at hente emissionsdata
def fetch_emissions_data(registration_number, api_token):
    """
    Henter emissionsdata fra API'et baseret på nummerpladen.
    """
    url = f"https://api.nrpla.de/emissions/{registration_number}"
    headers = {"Authorization": f"Bearer {api_token}"}
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()  # Tjek for HTTP-fejl
        return response.json()["data"].get("co2", None)  # Returnér CO2 værdien eller None
    except requests.exceptions.RequestException as e:
        raise Exception(f"API request failed: {e}")
    except Exception as e:
        raise Exception(f"Error fetching emissions data: {e}")


# Funktion til at hente fuel_type og fuel_efficiency hvis CO2 ikke er tilgængeligt
def fetch_fuel_data(registration_number, api_token):
    """
    Henter fuel_type og fuel_efficiency fra API'et baseret på nummerpladen,
    hvis CO2-udslippet ikke er tilgængeligt.
    """
    url = f"https://api.nrpla.de/{registration_number}"
    headers = {"Authorization": f"Bearer {api_token}"}
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()  # Tjek for HTTP-fejl
        data = response.json().get("data", {})

        if "fuel_type" not in data or "fuel_efficiency" not in data:
            raise Exception("Fuel type or fuel efficiency not available in API response.")

        fuel_type = data["fuel_type"]
        fuel_efficiency = str(data["fuel_efficiency"]).replace(".", ",")  # Erstat . med ,

        return fuel_type, fuel_efficiency

    except requests.exceptions.RequestException as e:
        raise Exception(f"API request failed: {e}")
    except Exception as e:
        raise Exception(f"Error fetching fuel data: {e}")


# Funktion til at beregne bilens alder
def calculate_vehicle_age(registration_date):
    """
    Beregner bilens alder baseret på registreringsdato og dags dato.
    """
    current_date = datetime.now()
    reg_date = datetime.strptime(registration_date, "%Y-%m-%d")
    age_years = (current_date - reg_date).days // 365
    return age_years


# Funktion til at finde handelsprisen baseret på alderen i Excel
def find_trade_price_based_on_age(file_path, vehicle_age):
    """
    Finder den handelspris, der skal bruges baseret på bilens alder og data i Excel-filen.
    """
    wb = load_workbook(file_path, data_only=True)  # Aktivér data-only for at evaluere formler
    sheet = wb["Ark1"]

    # Handelspris til regnemaskine findes i række 19 (kolonner E-I)
    if vehicle_age < 1:  # 0-1 år
        trade_price = sheet["E19"].value
        age_group = "0-1 år"
    elif 1 <= vehicle_age < 2:  # 1-2 år
        trade_price = sheet["F19"].value
        age_group = "1-2 år"
    elif 2 <= vehicle_age < 3:  # 2-3 år
        trade_price = sheet["G19"].value
        age_group = "2-3 år"
    elif 3 <= vehicle_age < 10:  # 3-9 år
        trade_price = sheet["H19"].value
        age_group = "3-9 år"
    else:  # Over 10 år
        trade_price = sheet["I19"].value
        age_group = "Over 10 år"

    return trade_price, age_group


# Funktion til at opdatere CO2 og fuel type i Excel-arket
def update_co2_in_excel(file_path, co2_value, fuel_type, fuel_efficiency):
    """
    Opdaterer CO2 og fuel type i Excel-arket 'Værktøj til CO2', samt opdaterer CO2 værdien i arket 'Brugte personvogne'.
    """
    app = xw.App(visible=False)
    wb = app.books.open(file_path)

    # Opdater 'Værktøj til CO2'
    sheet_co2_tool = wb.sheets["Værktøj til CO2"]

    # Sæt C26 til "NEDC"
    sheet_co2_tool.range("C26").value = "NEDC"

    if co2_value is None:
        # Hvis CO2 ikke er tilgængelig, udfør beregning baseret på fuel type og fuel efficiency
        sheet_co2_tool.range("C27").value = fuel_type  # Fuel type
        sheet_co2_tool.range("C25").value = fuel_efficiency  # Fuel efficiency
        co2_calculation = float(fuel_efficiency.replace(",", ".")) * 0.1  # Eksempel på formel for CO2 beregning
        sheet_co2_tool.range("C30").value = co2_calculation  # CO2 result i C30
    else:
        # Hvis CO2 er tilgængelig, opdater blot C30 med CO2 værdien
        sheet_co2_tool.range("C30").value = co2_value

    # Opdater 'Brugte personvogne' med CO2-værdien i L23
    sheet_trade_price = wb.sheets["Brugte personvogne"]

    # Hvis CO2-værdien er tilgængelig fra API, brug den værdi
    if co2_value is not None:
        sheet_trade_price.range("L23").value = co2_value  # Opdater L23 med CO2-værdien fra API
    else:
        sheet_trade_price.range("L23").value = sheet_co2_tool.range("C30").value  # Hvis beregning, brug C30's værdi

    wb.save()
    wb.close()
    app.quit()

    print(
        f"CO2 og fuel type er opdateret i 'Værktøj til CO2'. CO2: {co2_value}, Fuel type: {fuel_type}, Fuel efficiency: {fuel_efficiency}.")
    print(f"CO2 resultat i C30 er nu opdateret i 'Brugte personvogne' i celle L23.")


# Funktion til at opdatere Handelspris, Norm km og Kørte km i Excel-filen
def update_km_data(file_path, handelspris, norm_km, current_km):
    """
    Opdaterer Handelspris, Norm km og Kørte km i Excel-arket.
    """
    app = xw.App(visible=False)
    wb = app.books.open(file_path)
    sheet = wb.sheets["Ark1"]

    sheet.range("E7").value = handelspris
    sheet.range("E8").value = norm_km
    sheet.range("E9").value = current_km

    wb.save()
    wb.close()
    app.quit()

    print(f"Handelspris ({handelspris}), Norm km ({norm_km}) og Kørte km ({current_km}) er opdateret i '{file_path}'.")


# Funktion til at opdatere nyprisen og handelsprisen i arket 'Kopi-af-Bilafgifter-2021-v2.1-kopi'
def update_new_and_trade_price(file_path, new_price, trade_price):
    """
    Opdaterer nyprisen og handelsprisen i arket 'Kopi-af-Bilafgifter-2021-v2.1-kopi' i Excel-filen.
    """
    app = xw.App(visible=False)
    wb = app.books.open(file_path)
    sheet = wb.sheets["Brugte Personbiler"]

    sheet.range("L22").value = new_price  # Nypris sættes i L22
    sheet.range("L21").value = trade_price  # Handelspris sættes i L21

    wb.save()
    wb.close()
    app.quit()

    print(
        f"Nypris ({new_price}) og Handelspris ({trade_price}) er opdateret i 'Kopi-af-Bilafgifter-2021-v2.1-kopi.xlsx'.")


# Funktion til at beregne nypris
def fetch_new_price_from_api(registration_number, api_token):
    """
    Henter nyprisen for køretøjet fra API'et.
    Hvis 'retail_price' findes, bruges den. Ellers beregnes den som sum af 'evaluation' og 'Registration_tax'.
    """
    url = f"https://api.nrpla.de/evaluations/{registration_number}"
    headers = {"Authorization": f"Bearer {api_token}"}
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()  # Tjek for HTTP-fejl
        data = response.json()["data"][0]  # Første datasæt

        # Hvis retail_price findes, bruges den
        retail_price = data.get("retail_price")
        if retail_price is not None:
            new_price = retail_price
        else:
            # Beregn nypris som summen af evaluation og trading_price, hvis retail_price ikke findes
            new_price = data.get("evaluation", 0) + data.get("registration_tax", 0)

        return new_price

    except requests.exceptions.RequestException as e:
        raise Exception(f"API request failed: {e}")
    except Exception as e:
        raise Exception(f"Error fetching new price data: {e}")


# Funktion til at printe resultatet af G32 i 'Brugte personbiler'
def print_g32_value(file_path):
    """
    Henter og printer værdien af G32 fra arket 'Brugte personbiler'.
    """
    app = xw.App(visible=False)
    wb = app.books.open(file_path)
    sheet = wb.sheets["Brugte personbiler"]

    # Hent værdien fra G32
    g32_value = sheet.range("G32").value

    print(f"Eksportafgift: {g32_value}")

    wb.close()
    app.quit()

# Funktion til at hente nyprisen med fallback til manuel indtastning
def fetch_new_price_with_fallback(registration_number, api_token):
    """
    Henter nyprisen for køretøjet fra API'et.
    Hvis 'retail_price' findes, bruges den. Ellers beregnes den som sum af 'evaluation' og 'Registration_tax'.
    Hvis fejlen opstår, tillader den manuel indtastning af nyprisen.
    """
    try:
        # Forsøg at hente nyprisen fra API'et
        new_price = fetch_new_price_from_api(registration_number, api_token)
        print(f"Nypris for køretøjet: {new_price} kr.")
        return new_price
    except Exception as e:
        # Hvis der opstår en fejl, lad brugeren indtaste nyprisen manuelt
        print(f"Fejl ved hentning af nypris: {e}")
        new_price = float(input("Indtast nyprisen manuelt: "))
        print(f"Nypris indtastet manuelt: {new_price} kr.")
        return new_price


# Hovedfunktion (main) ændret til at bruge fallback for nypris
def main():
    registration_number = input("Indtast nummerplade: ")
    api_token = "g2Snwodm2Is5kaII6DpCNacdE9NCXF1vaUO9S2LG5S31v7svCT6JTfleb4FjInUj"
    km_file_path = "Kopi-af-Kopi-af-Km-beregning.xlsx"
    new_and_trade_file_path = "Kopi-af-Bilafgifter-2021-v2.1-kopi.xlsx"

    try:
        print("Henter køretøjsdata...")
        vehicle_data = fetch_vehicle_data(registration_number, api_token)
        print("Køretøjsdata hentet succesfuldt.")

        print("Henter emissionsdata...")
        co2_value = fetch_emissions_data(registration_number, api_token)
        print(f"Emissionsdata hentet: CO2: {co2_value}")

        registration_date = vehicle_data["date"]
        vehicle_age = calculate_vehicle_age(registration_date)
        print(f"Bilens alder: {vehicle_age} år")

        # Hvis CO2 ikke findes, hent fuel_type og fuel_efficiency
        if co2_value is None:
            print("CO2 ikke tilgængeligt, henter fuel_type og fuel_efficiency...")
            fuel_type, fuel_efficiency = fetch_fuel_data(registration_number, api_token)
        else:
            fuel_type, fuel_efficiency = None, None

        print(f"Fuel type: {fuel_type}, Fuel efficiency: {fuel_efficiency}")

        # Først, indtast handelspris, norm km og kørte km
        handelspris_input = float(input("Indtast handelsprisen: "))
        norm_km_input = float(input("Indtast norm km: "))
        current_km_input = float(input("Indtast bilens kørte kilometer: "))

        # Opdater værdierne i Excel
        update_km_data(km_file_path, handelspris_input, norm_km_input, current_km_input)

        # Beregn og få handelsprisen baseret på bilens alder fra Excel
        handelspris, age_group = find_trade_price_based_on_age(km_file_path, vehicle_age)
        print(f"Handelspris fra Excel: {handelspris} kr. for aldersgruppen {age_group}.")

        # Brug fallback til at hente nypris
        new_price = fetch_new_price_with_fallback(registration_number, api_token)

        # Opdater CO2 og fuel type i Excel
        update_co2_in_excel(new_and_trade_file_path, co2_value, fuel_type, fuel_efficiency)

        # Opdater handelsprisen baseret på alder i L21 og nyprisen i L22 i 'Kopi-af-Bilafgifter-2021-v2.1-kopi.xlsx'
        update_new_and_trade_price(new_and_trade_file_path, new_price, handelspris)

        print("\nProcessen er færdig!")

        # Hent og print resultatet af G32 i 'Brugte personbiler'
        print_g32_value(new_and_trade_file_path)

    except Exception as e:
        print(f"Fejl: {e}")


if __name__ == "__main__":
    main()