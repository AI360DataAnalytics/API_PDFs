from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel
from typing import Dict, Any

import matplotlib.pyplot as plt
import plotly.graph_objects as go
import numpy as np
import pandas as pd
import random
import string
import os
from docxtpl import DocxTemplate
import requests
import comtypes.client
import docx
import boto3

class Appraisal(BaseModel):
    id_apprasial: str
    id_apprasial_ai360:str

    street: str
    block: str
    zip_code: str
    locality: str
    state: str

    lat: str
    lng: str

    type: int
    land_surface: float
    built_surface: float
    age: int
    rooms: int
    bathrooms: int
    parking_slots: int
    warehouse: int
    finishes: int
    amenities: int
    roof_garden: int
    balcony: int
    outside_view: int
    level: int
    flats: int
    estimate_value_client: float

def set_pad_for_column(tbl, col, pad=0.1):
    cells = tbl.get_celld()
    column = [cell for cell in tbl.get_celld() if cell[1] == col]
    for cell in column:
        cells[cell].PAD = pad

def clean_files(files_to_delete):
    for f in files_to_delete:
        os.remove(f)

def render_mpl_table(data, col_width=3.0, row_height=0.625, font_size=14, cell_loc = 'center', header_color='#40466e', 
                     row_colors=['#f1f1f2', 'w'], edge_color='w', bbox=[0, 0, 1, 1], header_columns=0, **kwargs):
    # if ax is None:
    size = (np.array(data.shape[::-1]) + np.array([0, 1])) * np.array([col_width, row_height])
    fig, ax = plt.subplots(figsize=size)
    ax.axis('off')

    mpl_table = ax.table(cellText=data.values, bbox=bbox, colLabels=data.columns, cellLoc=cell_loc, **kwargs)
    mpl_table.auto_set_font_size(False)
    mpl_table.set_fontsize(font_size)
    mpl_table.auto_set_column_width(col=list(range(len(data.columns)))) # Provide integer list of columns to adjust

    for k, cell in mpl_table._cells.items():
        cell.set_edgecolor(edge_color)
        if k[0] == 0 or k[1] < header_columns:
            cell.set_text_props(weight='bold', color='w')
            cell.set_facecolor(header_color)
        else:
            cell.set_facecolor(row_colors[k[0]%len(row_colors)])

    return mpl_table

def table_precios(updated_min_total, updated_min_m2, updated_estimate_total, updated_estimate_m2, updated_max_total, updated_max_m2):
    
    df = pd.DataFrame()
    df['Precios AI360'] = []
    df['Valor Total'] = []
    df['Valor M2 Venta'] = []

    df.loc[0] = ['Mínimo','$ ' + format(updated_min_total, ","),'$ ' + format(updated_min_m2, ",")]
    df.loc[1] = ['Estimado','$ ' + format(updated_estimate_total, ","),'$ ' + format(updated_estimate_m2, ",")]
    df.loc[2] = ['Máximo','$ ' + format(updated_max_total, ","),'$ ' + format(updated_max_m2, ",")]

    mpl_table = render_mpl_table(df, header_columns=0, col_width=3.0)
    # look & feel adjustment
    for k, cell in mpl_table._cells.items():
        if k[0] == 2:
            cell.set_text_props(weight='bold')
            cell.get_text().set_fontsize(18)

    img_name = ''.join(random.choices(string.ascii_letters + string.digits, k=12))
    img_path = f'static/reports/sections/{img_name}.png'
    mpl_table.get_figure().savefig(img_path, bbox_inches = 'tight', pad_inches = 0)

    return img_path

def table_stats(stats):
    
    stat_count = int(float(stats[0]))
    stat_avg = int(float(stats[1]))
    stat_min = int(float(stats[2]))
    stat_q10 = int(float(stats[3]))
    stat_q20 = int(float(stats[4]))
    stat_q30 = int(float(stats[5]))
    stat_q40 = int(float(stats[6]))
    stat_q50 = int(float(stats[7]))
    stat_q60 = int(float(stats[8]))
    stat_q70 = int(float(stats[9]))
    stat_q80 = int(float(stats[10]))
    stat_q90 = int(float(stats[11]))
    stat_max = int(float(stats[12]))

    df = pd.DataFrame()
    df['Transacciones'] = []
    df['Promedio'] = []
    df['Min'] = []
    df['10%'] = []
    df['20%'] = []
    df['30%'] = []
    df['40%'] = []
    df['50%'] = []
    df['60%'] = []
    df['70%'] = []
    df['80%'] = []
    df['90%'] = []
    df['Max'] = []

    stat_count_lbl = format(stat_count, ",")
    stat_avg_lbl = '$ ' + format(stat_avg, ",")
    stat_min_lbl = '$ ' + format(stat_min, ",")
    stat_q10_lbl = '$ ' + format(stat_q10, ",")
    stat_q20_lbl = '$ ' + format(stat_q20, ",")
    stat_q30_lbl = '$ ' + format(stat_q30, ",")
    stat_q40_lbl = '$ ' + format( stat_q40, ",")
    stat_q50_lbl = '$ ' + format(stat_q50, ",")
    stat_q60_lbl = '$ ' + format(stat_q60, ",")
    stat_q70_lbl = '$ ' + format(stat_q70, ",")
    stat_q80_lbl = '$ ' + format(stat_q80, ",")
    stat_q90_lbl = '$ ' + format(stat_q90, ",")
    stat_max_lbl = '$ ' + format(stat_max, ",")

    df.loc[0] = [stat_count_lbl, stat_avg_lbl, stat_min_lbl, stat_q10_lbl, stat_q20_lbl, stat_q30_lbl, 
                    stat_q40_lbl, stat_q50_lbl, stat_q60_lbl, stat_q70_lbl, stat_q80_lbl, stat_q90_lbl, 
                    stat_max_lbl]

    mpl_table = render_mpl_table(df, header_columns=0, col_width=1.2)

    img_name = ''.join(random.choices(string.ascii_letters + string.digits, k=12))
    img_path = f'static/reports/sections/{img_name}.png'
    mpl_table.get_figure().savefig(img_path, bbox_inches = 'tight', pad_inches = 0)
    
    return img_path
    
def table_comparables(comparables, latitud, longitud):

    origin_marker_color = '#f0462f'
    alphabet_ids = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']
    neighbour_marker_color = '#21a8fa'
    style = 'streets'

    if comparables != '':

        df = pd.DataFrame()
        df['ID'] = []
        df['CP'] = []
        df['Tipo'] = []
        df['M2 Construído'] = []
        df['Precio M2'] = []
        df['Precio Total'] = []
        df['URL'] = []

        i = 0
        urls = []
        markers_ids = ['O']
        markers_colors = [origin_marker_color]
        symbols = ['star']
        address_lat = [float(latitud)]
        address_lng = [float(longitud)]
        
        for offer in comparables:
            if i < 10:
                alphabet_id = alphabet_ids[i]
                zip_code = f"{offer['zip_code']}"
                prop_type = f"{offer['type']}"
                built_surface = format(int(float(offer['built_surface']))) if offer['built_surface'] is not None else ''
                price_m2 = '$ ' + format(int(float(offer['price_m2'])), ",")
                price = '$ ' + format(int(float(offer['price'])), ",")

                row = [alphabet_id, zip_code, prop_type, built_surface, price_m2, price]

                url = offer['url_ad'] if offer['url_ad'] is not None else ''
                row.append('')
                urls.append(url)

                df.loc[i] = row

                if offer['lat'] is not None and offer['lng'] is not None:
                    neighbour_lat = float(offer['lat'])
                    neighbour_lng = float(offer['lng'])

                    address_lat.append(neighbour_lat)
                    address_lng.append(neighbour_lng)
                    markers_ids.append(alphabet_id)
                    markers_colors.append(neighbour_marker_color)
                    symbols.append('circle')

                i += 1
            else:
                break

        ############ Neighbours Map ##########################################################################
        map_data = {
                        'name': markers_ids,
                        'latitude': address_lat, 
                        'longitude': address_lng 
                    } 

        df_n = pd.DataFrame(data=map_data)
        

        map_zoom = 12

        map_data = go.Scattermapbox(lat=list(df_n['latitude']),
                                    lon=list(df_n['longitude']),
                                    mode='markers+text',
                                    marker = dict(size=12, color=markers_colors),
                                    textposition='top right',
                                    textfont=dict(size=16, color='black'),
                                    text=[df_n['name'][i] for i in range(df_n.shape[0])])

        map_layout = dict(margin=dict(l=0, t=0, r=0, b=0, pad=0),
                            mapbox=dict(accesstoken="pk.eyJ1IjoiYWkzNjAiLCJhIjoiY2tyY2Zzemx4MDdxdTJvbDkwNGQ1Z2Q4eCJ9.Ys6vnjreyN22ygvy9L-StA",
                            center=dict(lat=df_n['latitude'][0], lon=df_n['longitude'][0]),
                            style=style,
                            zoom=map_zoom))

        fig = go.Figure(data=map_data, layout=map_layout)

        img_name = ''.join(random.choices(string.ascii_letters + string.digits, k=12))
        local_ng_image_path = f'static/reports/sections/{img_name}.png'
        fig.write_image(local_ng_image_path)

        ############ Neighbours Table #########################################################################
        mpl_table = render_mpl_table(df, header_columns=0, col_width=3.0)

        img_name = ''.join(random.choices(string.ascii_letters + string.digits, k=12))
        img_path = f'static/reports/sections/{img_name}.png'
        mpl_table.get_figure().savefig(img_path, bbox_inches = 'tight', pad_inches = 0)

    return local_ng_image_path, img_path
        
def graph_prices(stats, updated_min_m2, updated_estimate_m2, updated_max_m2):

    stat_min = int(float(stats[2]))
    stat_q10 = int(float(stats[3]))
    stat_q20 = int(float(stats[4]))
    stat_q30 = int(float(stats[5]))
    stat_q40 = int(float(stats[6]))
    stat_q50 = int(float(stats[7]))
    stat_q60 = int(float(stats[8]))
    stat_q70 = int(float(stats[9]))
    stat_q80 = int(float(stats[10]))
    stat_q90 = int(float(stats[11]))
    stat_max = int(float(stats[12]))
    percentiles = np.array([stat_min, stat_q10, stat_q20, stat_q30, stat_q40, stat_q50, stat_q60, stat_q70, 
                        stat_q80, stat_q90, stat_max])

    plt.close('all')
    plt.clf()
    plt.grid()

    plt.plot([int(updated_min_m2), int(updated_min_m2)], [0, 10],  linestyle = '--', color = 'darkorange')
    plt.plot([int(updated_estimate_m2), int(updated_estimate_m2)], [0, 10], linestyle = '--', color = 'g')
    plt.plot([int(updated_max_m2), int(updated_max_m2)], [0, 10], linestyle = '--', color = 'darkorange')
    plt.plot(percentiles, ['min', '10%', '20%', '30%', '40%', '50%', '60%', '70%', '80%', '90%', 'max'])

    img_name = ''.join(random.choices(string.ascii_letters + string.digits, k=12))
    img_path = f'static/reports/sections/{img_name}.png'
    plt.savefig(img_path)

    return img_path

def table_score(raitings):
    star_char = u'\u2605'

    r_nse = raitings['nse']
    r_commercialization = raitings['commercialization']
    r_capital_gain = raitings['capital_gain']
    r_amenities = raitings['amenities']
    r_price = raitings['price']
    r_time_on_market = raitings['time_on_market']
    r_global = raitings['global']

    rating_nse = round(float(r_nse) * .05)
    rating_comm = round(float(r_commercialization) * .05)
    rating_capital = round(float(r_capital_gain) * .05)
    rating_amenities = round(float(r_amenities) * .05)
    rating_price = round(float(r_price) * .05)
    rating_time_on_market =round(float(r_time_on_market) * .05)
    rating_global = round(float(r_global) *.05, 1)
    rating_global_rounded = round(float(rating_global))

    rating_nse_stars = ' '.join(star_char * rating_nse)
    rating_comm_stars = ' '.join(star_char * rating_comm)
    rating_capital_stars = ' '.join(star_char * rating_capital)
    rating_amenities_stars = ' '.join(star_char * rating_amenities)
    rating_price_stars = ' '.join(star_char * rating_price)
    rating_global_stars = ' '.join(star_char * rating_global_rounded)
    rating_time_on_market_stars = ' '.join(star_char * rating_time_on_market)

    ########## Scale Clasification ########################################################################
    df = pd.DataFrame()
    df['Características'] = []
    df['Interpretación'] = []
    df['Puntuación'] = []
    df['Estrellas'] = []

    df.loc[0] = ['Nivel Socioeconómico', '¿Qué tan alto es el nivel de ingreso de la manzana respecto al municipio?', r_nse, rating_nse_stars]
    df.loc[1] = ['Comercialización', '¿Cuántas transacciones se hacen cen el CP respecto al municipio?', r_commercialization, rating_comm_stars]
    df.loc[2] = ['Plusvalía', '¿Qué tanto crecen los precios del CP vs los del municipio?', r_capital_gain, rating_capital_stars]
    df.loc[3] = ['Servicios', '¿Cuántos servicios se ubican en el CP respecto del total del municipio?', r_amenities, rating_amenities_stars]
    df.loc[4] = ['Precios', '¿Qué tan alto o bajo es el precio estimado vs el municipio?', r_price, rating_price_stars]
    df.loc[5] = ['Tiempo de venta', '¿Qué tan rapido se puede vender respecto a los tiempos de venta del municipio?', r_time_on_market, rating_time_on_market_stars]
    df.loc[6] = ['Calificación global', '¿Cuál es la calidad de esta garantía?', r_global, rating_global_stars]

    mpl_table = render_mpl_table(df, header_columns=0, col_width=5.0, font_size=14)
    # look & feel adjustment
    for k, cell in mpl_table._cells.items():
        if k[0] > 0 and k[1] == 3:
            cell.get_text().set_color('#ffcd01')
            cell.get_text().set_fontsize(28)

    set_pad_for_column(mpl_table, 2, pad=0.46)

    img_name = ''.join(random.choices(string.ascii_letters + string.digits, k=12))
    img_path = f'static/reports/sections/{img_name}.png'
    mpl_table.get_figure().savefig(img_path, bbox_inches = 'tight', pad_inches = 0)

    return img_path

def obtain_street_view_image(street, block2, municipality, state):
    img_name = None

    base_url =  "https://maps.googleapis.com/maps/api/streetview"
    api_key = "AIzaSyAMSZbv76mXJxCHkGZVPJV2wMFmLH0W1mE"

    address = f'{street},{block2},{municipality},{state}'
    address = address.replace('#', '')

    url = f"{base_url}?size=1200x800&location='{address}'&key={api_key}"

    r = requests.get(url)

    if r.status_code == 200:
        img_name = ''.join(random.choices(string.ascii_letters + string.digits, k=12))

        f=open(f'static/reports/sections/{img_name}.jpg','wb')
        img_path = f'static/reports/sections/{img_name}.jpg'
        f.write(r.content)
        f.close()


    return img_path

def insert_bucket(pdf_path, id_apprasail):

    session = boto3.Session(
        aws_access_key_id= "AKIAU32QZGPMBXMGGCPV",
        aws_secret_access_key="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"
    )
    s3 = session.resource('s3')

    s3.meta.client.upload_file(Filename=pdf_path, Bucket='ai360-bpb-testpdf', Key = f'{id_apprasail}.pdf')

    return 0

def transform_to_PDF(doc_path, pdf_path):

    word = comtypes.client.CreateObject("Word.Application")
    docx_path = os.path.abspath(doc_path)
    pdf_path = os.path.abspath(pdf_path)

    pdf_format = 17
    word.Visible = False
    in_file = word.Documents.Open(docx_path)
    in_file.SaveAs(pdf_path, FileFormat = pdf_format)
    in_file.Close()

    word.Quit()

    return 0

router = APIRouter()

@router.post("/pdf_BPB_json")
async def estimate_appraisal(appraisal: Dict[str, Any] = {
    "property": {
        "requestor_details": {
            "id_apprasial": "29244",
            "id_apprasial_ai360": "bpb-db67560d-40ee-48be-a438-ab81347db87c"
        },
        "address": {
            "street": "av periferico 180",
            "block": "C.T.M. Atzacoalco",
            "zip_code": "07090",
            "locality": "Gustavo A. Madero",
            "state": "Ciudad de M\u00e9xico"
        },
        "geolocation": {
            "lat": "19.50727760",
            "lng": "-99.10084010"
        },
        "info": {
            "type": 2,
            "land_surface": "150.00",
            "built_surface": "170.00",
            "age": 30,
            "rooms": 4,
            "bathrooms": 2,
            "parking_slots": 2,
            "warehouse": 1,
            "finishes": 0,
            "amenities": 0,
            "roof_garden": 0,
            "balcony": 1,
            "outside_view": 1,
            "level": 0,
            "flats": 2,
            "estimate_value_client": 2100000
        },
        "apprasial": {
            "conservation": "Usado",
            "class": "Interes social",
            "time_on_market": "3.1",
            "updated_estimate_total": "2941000",
            "updated_estimate_m2": "17303",
            "updated_min_total": "2499000",
            "updated_min_m2": "14707",
            "updated_max_total": "3382000",
            "updated_max_m2": "19898",
            "raitings": {
                "amenities": "100",
                "capital_gain": "60",
                "commercialization": "100",
                "global": "70",
                "nse": "40",
                "price": "60",
                "time_on_market": "60"
            },
            "statistics_m2": {
                "count": "6",
                "mean": "16603",
                "min": "12494",
                "q10": "14148",
                "q20": "15803",
                "q30": "16295",
                "q40": "16788",
                "q50": "16918",
                "q60": "17048",
                "q70": "17158",
                "q80": "17268",
                "q90": "18744",
                "max": "20220"
            },
            "similar_properties": [
                {
                    "type": "Casa",
                    "lat": "19.5094027",
                    "lng": "-99.1020413",
                    "zip_code": "7090",
                    "price": "3200000",
                    "price_m2": "19277",
                    "built_surface": "166",
                    "rooms": "3",
                    "bathrooms": "2",
                    "parking_slots": "2",
                    "homologation_factor": "",
                    "distance": "0.27",
                    "url_ad": "https://propiedades.com/inmuebles/casa-en-venta-retorno-jesus-aldrete-11-ctm-atzacoalco-df-19505601#pagina=12&tipos=casas-venta&area=df&pos=36",
                    "url_images": "https://propiedadescom.s3.amazonaws.com/files/600x400/6374426901904c68a8b97e403b6d034c.png, https://propiedadescom.s3.amazonaws.com/files/600x400/25e358f13b4b871a5a7c2f9e3399aee0.jpeg, https://propiedadescom.s3.amazonaws.com/files/600x400/9865f98c3a2882166ad22dc9d5b9c754.jpeg, https://propiedadescom.s3.amazonaws.com/files/600x400/e7a4b67ee289237c77c1bbeb04f65ea7.jpeg, https://propiedadescom.s3.amazonaws.com/files/600x400/5d3d63353aa670a55712552409df2d70.jpeg"
                },
                {
                    "type": "Casa",
                    "lat": "19.5073993",
                    "lng": "-99.097238",
                    "zip_code": "07090",
                    "price": "2950000",
                    "price_m2": "16388",
                    "built_surface": "180",
                    "rooms": "4",
                    "bathrooms": "2",
                    "parking_slots": "1",
                    "homologation_factor": "",
                    "distance": "0.38",
                    "url_ad": "http://www.inmuebles24.com/propiedades/retorno-maximo-molina-unidad-ctm-atzacoalco.-gustavo-60979179.html",
                    "url_images": "https://img10.naventcdn.com/avisos/resize/18/00/60/97/91/79/1200x1200/301015932.jpg,https://img10.naventcdn.com/avisos/resize/18/00/60/97/91/79/1200x1200/301015784.jpg,https://img10.naventcdn.com/avisos/resize/18/00/60/97/91/79/1200x1200/301015630.jpg,https://img10.naventcdn.com/avisos/resize/18/00/60/97/91/79/1200x1200/301015638.jpg,https://img10.naventcdn.com/avisos/resize/18/00/60/97/91/79/1200x1200/301015628.jpg,https://img10.naventcdn.com/avisos/resize/18/00/60/97/91/79/1200x1200/301015687.jpg,https://img10.naventcdn.com/avisos/resize/18/00/60/97/91/79/1200x1200/301015625.jpg,https://img10.naventcdn.com/avisos/resize/18/00/60/97/91/79/1200x1200/301015786.jpg,"
                },
                {
                    "type": "Casa",
                    "lat": "19.5074",
                    "lng": "-99.09724",
                    "zip_code": "07090",
                    "price": "3000000",
                    "price_m2": "16666",
                    "built_surface": "180",
                    "rooms": "5",
                    "bathrooms": "2",
                    "parking_slots": "1",
                    "homologation_factor": "",
                    "distance": "0.38",
                    "url_ad": "http://www.inmuebles24.com/propiedades/casa-en-venta-en-c.-t.-m.-atzacoalco-62295810.html",
                    "url_images": "https://img10.naventcdn.com/avisos/resize/18/00/62/29/58/10/1200x1200/280205190.jpg,https://img10.naventcdn.com/avisos/resize/18/00/62/29/58/10/1200x1200/280205194.jpg,https://img10.naventcdn.com/avisos/resize/18/00/62/29/58/10/1200x1200/280205189.jpg,https://img10.naventcdn.com/avisos/resize/18/00/62/29/58/10/1200x1200/280205188.jpg,https://img10.naventcdn.com/avisos/resize/18/00/62/29/58/10/1200x1200/314405795.jpg,https://img10.naventcdn.com/avisos/resize/18/00/62/29/58/10/1200x1200/321124999.jpg,https://img10.naventcdn.com/avisos/resize/18/00/62/29/58/10/1200x1200/280205182.jpg,https://img10.naventcdn.com/avisos/resize/18/00/62/29/58/10/1200x1200/280205192.jpg,"
                },
                {
                    "type": "Casa",
                    "lat": "19.506097",
                    "lng": "-99.093388",
                    "zip_code": "07090",
                    "price": "3500000",
                    "price_m2": "19886",
                    "built_surface": "176",
                    "rooms": "3",
                    "bathrooms": "3",
                    "parking_slots": "1",
                    "homologation_factor": "",
                    "distance": "0.79",
                    "url_ad": "https://propiedades.com/inmuebles/casa-en-venta-fernando-amilpa-239-ctm-atzacoalco-df-17699305#area=gustavo-a-madero&tipos=casas-venta&orden=&pagina=7&paginas=83&pos=1",
                    "url_images": "https://propiedadescom.s3.amazonaws.com/files/600x400/fernando-amilpa-239-ctm-atzacoalco-gustavo-a-madero-df-cdmx-17699305-foto-05.jpg, https://propiedadescom.s3.amazonaws.com/files/600x400/fernando-amilpa-239-ctm-atzacoalco-gustavo-a-madero-df-cdmx-17699305-foto-06.jpg, https://propiedadescom.s3.amazonaws.com/files/600x400/fernando-amilpa-239-ctm-atzacoalco-gustavo-a-madero-df-cdmx-17699305-foto-09.jpg"
                },
                {
                    "type": "Casa",
                    "lat": "19.5089693",
                    "lng": "-99.0914215",
                    "zip_code": "07090",
                    "price": "2750000",
                    "price_m2": "15277",
                    "built_surface": "180",
                    "rooms": "2",
                    "bathrooms": "4",
                    "parking_slots": "2",
                    "homologation_factor": "",
                    "distance": "1.01",
                    "url_ad": "https://www.inmuebles24.com/propiedades/venta-casa-ciudad-de-mexico-ctm-el-risco-58343346.html",
                    "url_images": "https://img10.naventcdn.com/avisos/resize/18/00/58/34/33/46/1200x1200/211723617.jpg, https://img10.naventcdn.com/avisos/resize/18/00/58/34/33/46/1200x1200/211720262.jpg, https://img10.naventcdn.com/avisos/resize/18/00/58/34/33/46/1200x1200/211720265.jpg, https://img10.naventcdn.com/avisos/resize/18/00/58/34/33/46/1200x1200/211720261.jpg, https://img10.naventcdn.com/avisos/resize/18/00/58/34/33/46/1200x1200/211720263.jpg, https://img10.naventcdn.com/avisos/resize/18/00/58/34/33/46/1200x1200/211720267.jpg, https://img10.naventcdn.com/avisos/resize/18/00/58/34/33/46/1200x1200/211720257.jpg, https://img10.naventcdn.com/avisos/resize/18/00/58/34/33/46/1200x1200/211720260.jpg, https://img10.naventcdn.com/avisos/resize/18/00/58/34/33/46/1200x1200/211720259.jpg, https://img10.naventcdn.com/avisos/resize/18/00/58/34/33/46/1200x1200/211720264.jpg, https://img10.naventcdn.com/avisos/resize/18/00/58/34/33/46/1200x1200/211720258.jpg, https://img10.naventcdn.com/avisos/resize/18/00/58/34/33/46/1200x1200/211720266.jpg"
                },
                {
                    "type": "Casa",
                    "lat": "19.5023775",
                    "lng": "-99.0897555",
                    "zip_code": "07410",
                    "price": "2900000",
                    "price_m2": "16111",
                    "built_surface": "180",
                    "rooms": "2",
                    "bathrooms": "4",
                    "parking_slots": "1",
                    "homologation_factor": "",
                    "distance": "1.28",
                    "url_ad": "https://propiedades.com/inmuebles/casa-en-venta-norte-74-a-ampliacion-emiliano-zapata-df-6125166#area=gustavo-a-madero&tipos=casas-venta&orden=&pagina=30&paginas=83&pos=1",
                    "url_images": "https://propiedadescom.s3.amazonaws.com/files/600x400/ac268a823e9e5058a18bd87615cf0bf0.png, https://propiedadescom.s3.amazonaws.com/files/600x400/fa8caad5c044707485695166decf1dd5.png, https://propiedadescom.s3.amazonaws.com/files/600x400/b4c88c964cdecf2a5e97393fa0c2c4a0.png, https://propiedadescom.s3.amazonaws.com/files/600x400/089d905868391fd3d3ff9c072b38546a.png"
                },
                {
                    "type": "Casa",
                    "lat": "19.4994338",
                    "lng": "-99.1119157",
                    "zip_code": "07010",
                    "price": "2600000",
                    "price_m2": "17333",
                    "built_surface": "150",
                    "rooms": "1",
                    "bathrooms": "3",
                    "parking_slots": "1",
                    "homologation_factor": "",
                    "distance": "1.45",
                    "url_ad": "https://www.inmuebles24.com/propiedades/casa-en-santa-isabel-tola-gustavo-a.-madero-58003610.html",
                    "url_images": "https://img10.naventcdn.com/avisos/resize/18/00/58/00/36/10/1200x1200/206453571.jpg, https://img10.naventcdn.com/avisos/resize/18/00/58/00/36/10/1200x1200/206453574.jpg, https://img10.naventcdn.com/avisos/resize/18/00/58/00/36/10/1200x1200/206453572.jpg, https://img10.naventcdn.com/avisos/resize/18/00/58/00/36/10/1200x1200/206453576.jpg, https://img10.naventcdn.com/avisos/resize/18/00/58/00/36/10/1200x1200/206453575.jpg, https://img10.naventcdn.com/avisos/resize/18/00/58/00/36/10/1200x1200/206453573.jpg"
                },
                {
                    "type": "Casa",
                    "lat": "19.4998798",
                    "lng": "-99.1123524",
                    "zip_code": "07010",
                    "price": "2600000",
                    "price_m2": "17333",
                    "built_surface": "150",
                    "rooms": "2",
                    "bathrooms": "3",
                    "parking_slots": "0",
                    "homologation_factor": "",
                    "distance": "1.46",
                    "url_ad": "https://www.inmuebles24.com/propiedades/casa-oportunidad-en-colonia-santa-isabel-tola-del-58306502.html",
                    "url_images": "0"
                },
                {
                    "type": "Casa",
                    "lat": "19.50052",
                    "lng": "-99.114573",
                    "zip_code": "07010",
                    "price": "3000000",
                    "price_m2": "18750",
                    "built_surface": "160",
                    "rooms": "2",
                    "bathrooms": "3",
                    "parking_slots": "2",
                    "homologation_factor": "",
                    "distance": "1.62",
                    "url_ad": "https://propiedades.com/inmuebles/casa-en-venta-de-los-faraones-9-acueducto-de-guadalupe-df-18162455#area=gustavo-a-madero&tipos=casas-venta&orden=&pagina=51&paginas=83&pos=1",
                    "url_images": "https://propiedadescom.s3.amazonaws.com/files/600x400/de-los-faraones-9-acueducto-de-guadalupe-gustavo-a-madero-df-cdmx-18162455-foto-01.jpg, https://propiedadescom.s3.amazonaws.com/files/600x400/de-los-faraones-9-acueducto-de-guadalupe-gustavo-a-madero-df-cdmx-18162455-foto-02.jpg, https://propiedadescom.s3.amazonaws.com/files/600x400/de-los-faraones-9-acueducto-de-guadalupe-gustavo-a-madero-df-cdmx-18162455-foto-03.jpg, https://propiedadescom.s3.amazonaws.com/files/600x400/de-los-faraones-9-acueducto-de-guadalupe-gustavo-a-madero-df-cdmx-18162455-foto-04.jpg, https://propiedadescom.s3.amazonaws.com/files/600x400/de-los-faraones-9-acueducto-de-guadalupe-gustavo-a-madero-df-cdmx-18162455-foto-05.jpg, https://propiedadescom.s3.amazonaws.com/files/600x400/de-los-faraones-9-acueducto-de-guadalupe-gustavo-a-madero-df-cdmx-18162455-foto-06.jpg, https://propiedadescom.s3.amazonaws.com/files/600x400/de-los-faraones-9-acueducto-de-guadalupe-gustavo-a-madero-df-cdmx-18162455-foto-07.jpg, https://propiedadescom.s3.amazonaws.com/files/600x400/de-los-faraones-9-acueducto-de-guadalupe-gustavo-a-madero-df-cdmx-18162455-foto-08.jpg, https://propiedadescom.s3.amazonaws.com/files/600x400/de-los-faraones-9-acueducto-de-guadalupe-gustavo-a-madero-df-cdmx-18162455-foto-09.jpg, https://propiedadescom.s3.amazonaws.com/files/600x400/de-los-faraones-9-acueducto-de-guadalupe-gustavo-a-madero-df-cdmx-18162455-foto-10.jpg, https://propiedadescom.s3.amazonaws.com/files/600x400/de-los-faraones-9-acueducto-de-guadalupe-gustavo-a-madero-df-cdmx-18162455-foto-11.jpg, https://propiedadescom.s3.amazonaws.com/files/600x400/de-los-faraones-9-acueducto-de-guadalupe-gustavo-a-madero-df-cdmx-18162455-foto-12.jpg, https://propiedadescom.s3.amazonaws.com/files/600x400/de-los-faraones-9-acueducto-de-guadalupe-gustavo-a-madero-df-cdmx-18162455-foto-13.jpg, https://propiedadescom.s3.amazonaws.com/files/600x400/de-los-faraones-9-acueducto-de-guadalupe-gustavo-a-madero-df-cdmx-18162455-foto-14.jpg, https://propiedadescom.s3.amazonaws.com/files/600x400/de-los-faraones-9-acueducto-de-guadalupe-gustavo-a-madero-df-cdmx-18162455-foto-15.jpg, https://propiedadescom.s3.amazonaws.com/files/600x400/de-los-faraones-9-acueducto-de-guadalupe-gustavo-a-madero-df-cdmx-18162455-foto-16.jpg, https://propiedadescom.s3.amazonaws.com/files/600x400/de-los-faraones-9-acueducto-de-guadalupe-gustavo-a-madero-df-cdmx-18162455-foto-17.jpg, https://propiedadescom.s3.amazonaws.com/files/600x400/de-los-faraones-9-acueducto-de-guadalupe-gustavo-a-madero-df-cdmx-18162455-foto-18.jpg, https://propiedadescom.s3.amazonaws.com/files/600x400/de-los-faraones-9-acueducto-de-guadalupe-gustavo-a-madero-df-cdmx-18162455-foto-19.jpg"
                },
                {
                    "type": "Casa",
                    "lat": "19.4959238",
                    "lng": "-99.0882817",
                    "zip_code": "07420",
                    "price": "2534000",
                    "price_m2": "14732",
                    "built_surface": "172",
                    "rooms": "2",
                    "bathrooms": "3",
                    "parking_slots": "2",
                    "homologation_factor": "",
                    "distance": "1.82",
                    "url_ad": "https://www.inmuebles24.com/propiedades/casa-de-remate-judicial-col.-nueva-atzacoalco-57082401.html",
                    "url_images": "https://img10.naventcdn.com/avisos/resize/18/00/57/08/24/01/1200x1200/140257723.jpg, https://img10.naventcdn.com/avisos/resize/18/00/57/08/24/01/1200x1200/140257720.jpg, https://img10.naventcdn.com/avisos/resize/18/00/57/08/24/01/1200x1200/140257721.jpg,"
                }
            ]
        }
        
    }
}):

    files_to_delete = list()

    id_apprasail = appraisal["property"]["requestor_details"]["id_apprasial"]

    age = appraisal["property"]["info"]["age"]
    land_surface = appraisal["property"]["info"]["land_surface"]
    built_surface = appraisal["property"]["info"]["built_surface"]
    rooms = appraisal["property"]["info"]["rooms"]
    bathrooms = appraisal["property"]["info"]["bathrooms"]
    parking_slots = appraisal["property"]["info"]["parking_slots"]

    updated_min_total = int(appraisal["property"]["apprasial"]["updated_min_total"])
    updated_estimate_total = int(appraisal["property"]["apprasial"]["updated_estimate_total"])
    updated_max_total = int(appraisal["property"]["apprasial"]["updated_max_total"])

    updated_min_m2 = int(appraisal["property"]["apprasial"]["updated_min_m2"])
    updated_estimate_m2 = int(appraisal["property"]["apprasial"]["updated_estimate_m2"])
    updated_max_m2 = int(appraisal["property"]["apprasial"]["updated_max_m2"])

    comparables =appraisal['property']['apprasial']['similar_properties']

    street = appraisal["property"]["address"]["street"]
    block = appraisal["property"]["address"]["block"]
    zip_code = appraisal["property"]["address"]["zip_code"]
    locality = appraisal["property"]["address"]["locality"]
    state = appraisal["property"]["address"]["state"]
    latitud = appraisal['property']['geolocation']['lat']
    longitud = appraisal['property']['geolocation']['lng']

    stats = appraisal["property"]['apprasial']['statistics_m2']
    stats = list(stats.values())

    raitings = appraisal['property']['apprasial']['raitings']

    path_table_precios = table_precios(updated_min_total, updated_min_m2, updated_estimate_total, updated_estimate_m2, updated_max_total, updated_max_m2)
    files_to_delete.append(path_table_precios)

    path_table_stats = table_stats(stats)
    files_to_delete.append(path_table_stats)

    path_table_comparables, path_map_comparables = table_comparables(comparables, latitud, longitud)
    files_to_delete.append(path_table_comparables)
    files_to_delete.append(path_map_comparables)

    path_graph_prices = graph_prices(stats, updated_min_m2, updated_estimate_m2, updated_max_m2)
    files_to_delete.append(path_graph_prices)

    path_table_score = table_score(raitings)
    files_to_delete.append(path_table_score)

    path_sv_image =obtain_street_view_image(street, block, locality, state)
    files_to_delete.append(path_sv_image)

    path_template = r"static\templates\Template_BPB.docx"

    doc = DocxTemplate(path_template)

    context = { 'id_apprasial' : id_apprasail,
            'state':state,
            'block': block, 
            'zip_code': zip_code,
            'locality': locality,
            'age': age,
            'land_surface': land_surface,
            "built_surface": built_surface,
            "parking_slots": parking_slots,
            "bathrooms": bathrooms,
            "rooms": rooms}

    #doc.replace_pic("img_sv",path_sv_image)
    doc.replace_pic("table_precios", path_table_precios)
    doc.replace_pic("table_stats", path_table_stats)
    doc.replace_pic("graph_prices", path_graph_prices)   
    doc.replace_pic("map_comparables", path_map_comparables) 
    # doc.replace_pic("table_comparables", path_table_comparables)
    doc.replace_pic("table_score", path_table_score)

    doc.render(context)

    doc_path = f"static\created\{id_apprasail}_generated_doc.docx"
    files_to_delete.append(doc_path)
    doc.save(doc_path)
        
    pdf_path = f"static\created\{id_apprasail}_generated_doc.pdf"

    doc = docx.Document(doc_path)

    word = comtypes.client.CreateObject("Word.Application")
    docx_path = os.path.abspath(doc_path)
    pdf_path = os.path.abspath(pdf_path)

    pdf_format = 17
    word.Visible = False
    in_file = word.Documents.Open(docx_path)
    in_file.SaveAs(pdf_path, FileFormat = pdf_format)
    in_file.Close()

    word.Quit()


    insert_bucket(pdf_path, id_apprasail)
   
    clean_files(files_to_delete)

    return files_to_delete