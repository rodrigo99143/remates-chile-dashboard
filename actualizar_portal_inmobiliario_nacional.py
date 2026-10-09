# -*- coding: utf-8 -*-
"""
BARRIDO NACIONAL DE PORTAL INMOBILIARIO (las 346 comunas de Chile)
====================================================================
Este es un script APARTE, pensado para correrlo de vez en cuando (por
ejemplo, una vez cada 6 meses) - NO se ejecuta solo cada vez que corres
Untitled-1.py. La idea (propuesta por Rodrigo) es:

  - Portal Inmobiliario es por lejos el sitio con más publicaciones de
    Chile, pero recorrer TODO el país (346 comunas x 2 tipos = ~690
    páginas) toma mucho rato. No tiene sentido hacer eso cada vez que
    corres el programa principal.
  - Los precios de propiedades no cambian de un día para otro, así que
    basta con actualizar este "catastro de precios por comuna" cada
    cierto tiempo (un semestre, por ejemplo), no todos los días.
  - Este script guarda sus resultados en la MISMA base de datos que ya
    usa precio_mercado.py (datos/precio_mercado.db). Una vez que termine
    de correr, Untitled-1.py lo va a usar automáticamente la próxima vez
    que lo corras - no hay que tocar nada más en el código principal.
  - Si el día de mañana Portal Inmobiliario cambia de sitio y deja de
    funcionar, igual queda guardado en la base de datos todo lo que ya se
    había recolectado - no se pierde, y el programa principal lo sigue
    usando con la fecha en que se recolectó cada publicación.

CÓMO USARLO:
  1. Guárdalo en la MISMA carpeta que Untitled-1.py y precio_mercado.py.
  2. Ábrelo en VS Code y presiona Run.
  3. Va a demorar harto (puede ser más de una hora: son ~690 páginas, con
     una pausa entre cada una para no sobrecargar el sitio). Puedes
     dejarlo corriendo de fondo y revisar más tarde.
  4. Si lo interrumpes a mitad de camino no pasa nada: lo que ya alcanzó
     a guardar queda guardado. Si lo vuelves a correr después, las
     comunas que ya había visto no se duplican (las marca "ya vistas" y
     sigue), así que puedes volver a correrlo entero sin miedo.
  5. Al final te muestra un resumen por región: cuántas comunas SÍ
     encontraron propiedades y cuántas quedaron en cero (esas últimas
     probablemente es porque el "slug" de esa región en la URL no es el
     que supusimos - mándame ese resumen y lo ajustamos puntualmente).

OJO - LAS REGIONES YA CONFIRMADAS (Biobío, Metropolitana, Valparaíso,
Coquimbo, Antofagasta, Araucanía, O'Higgins, Maule) se saltan acá porque
ya las cubre el barrido chico de todos los días (proveedor
"PortalInmobiliario" dentro de precio_mercado.py). Este script se
concentra en las OTRAS 8 regiones, que todavía no se habían probado
nunca (Arica y Parinacota, Tarapacá, Atacama, Ñuble, Los Ríos, Los
Lagos, Aysén, Magallanes) - ahí sí pueden fallar más URLs porque el
patrón de la región es una suposición todavía sin confirmar.
"""

import time
import datetime as dt
from pathlib import Path

import requests

import precio_mercado as pm

CARPETA_DATOS = Path(__file__).parent / "datos"
PAUSA_SEG = 1.5  # cortesía con el sitio - son muchas páginas


# Comunas YA cubiertas por el barrido chico de todos los días (no hace
# falta repetirlas acá, así se ahorra tiempo).
_YA_CUBIERTAS = {slug for slug, _nombre in pm._COMUNAS_OBJETIVO_PORTALINMOBILIARIO}


# Las 346 comunas de Chile, agrupadas por región. Cada tupla es:
# (slug para la URL, nombre de la comuna para la base de datos, región).
TODAS_LAS_COMUNAS_CHILE = [
    # ---- Arica y Parinacota ----
    ("arica-arica-parinacota", "ARICA", "Arica y Parinacota"),
    ("camarones-arica-parinacota", "CAMARONES", "Arica y Parinacota"),
    ("general-lagos-arica-parinacota", "GENERAL LAGOS", "Arica y Parinacota"),
    ("putre-arica-parinacota", "PUTRE", "Arica y Parinacota"),
    # ---- Tarapacá ----
    ("alto-hospicio-tarapaca", "ALTO HOSPICIO", "Tarapacá"),
    ("camina-tarapaca", "CAMINA", "Tarapacá"),
    ("colchane-tarapaca", "COLCHANE", "Tarapacá"),
    ("huara-tarapaca", "HUARA", "Tarapacá"),
    ("iquique-tarapaca", "IQUIQUE", "Tarapacá"),
    ("pica-tarapaca", "PICA", "Tarapacá"),
    ("pozo-almonte-tarapaca", "POZO ALMONTE", "Tarapacá"),
    # ---- Antofagasta ----
    ("antofagasta-antofagasta", "ANTOFAGASTA", "Antofagasta"),
    ("calama-antofagasta", "CALAMA", "Antofagasta"),
    ("maria-elena-antofagasta", "MARIA ELENA", "Antofagasta"),
    ("mejillones-antofagasta", "MEJILLONES", "Antofagasta"),
    ("ollague-antofagasta", "OLLAGUE", "Antofagasta"),
    ("san-pedro-de-atacama-antofagasta", "SAN PEDRO DE ATACAMA", "Antofagasta"),
    ("sierra-gorda-antofagasta", "SIERRA GORDA", "Antofagasta"),
    ("taltal-antofagasta", "TALTAL", "Antofagasta"),
    ("tocopilla-antofagasta", "TOCOPILLA", "Antofagasta"),
    # ---- Atacama ----
    ("alto-del-carmen-atacama", "ALTO DEL CARMEN", "Atacama"),
    ("caldera-atacama", "CALDERA", "Atacama"),
    ("chanaral-atacama", "CHANARAL", "Atacama"),
    ("copiapo-atacama", "COPIAPO", "Atacama"),
    ("diego-de-almagro-atacama", "DIEGO DE ALMAGRO", "Atacama"),
    ("freirina-atacama", "FREIRINA", "Atacama"),
    ("huasco-atacama", "HUASCO", "Atacama"),
    ("tierra-amarilla-atacama", "TIERRA AMARILLA", "Atacama"),
    ("vallenar-atacama", "VALLENAR", "Atacama"),
    # ---- Coquimbo ----
    ("andacollo-coquimbo", "ANDACOLLO", "Coquimbo"),
    ("canela-coquimbo", "CANELA", "Coquimbo"),
    ("combarbala-coquimbo", "COMBARBALA", "Coquimbo"),
    ("coquimbo-coquimbo", "COQUIMBO", "Coquimbo"),
    ("illapel-coquimbo", "ILLAPEL", "Coquimbo"),
    ("la-higuera-coquimbo", "LA HIGUERA", "Coquimbo"),
    ("la-serena-coquimbo", "LA SERENA", "Coquimbo"),
    ("los-vilos-coquimbo", "LOS VILOS", "Coquimbo"),
    ("monte-patria-coquimbo", "MONTE PATRIA", "Coquimbo"),
    ("ovalle-coquimbo", "OVALLE", "Coquimbo"),
    ("paiguano-coquimbo", "PAIGUANO", "Coquimbo"),
    ("punitaqui-coquimbo", "PUNITAQUI", "Coquimbo"),
    ("rio-hurtado-coquimbo", "RIO HURTADO", "Coquimbo"),
    ("salamanca-coquimbo", "SALAMANCA", "Coquimbo"),
    ("vicuna-coquimbo", "VICUNA", "Coquimbo"),
    # ---- Valparaíso ----
    ("algarrobo-valparaiso", "ALGARROBO", "Valparaíso"),
    ("cabildo-valparaiso", "CABILDO", "Valparaíso"),
    ("calle-larga-valparaiso", "CALLE LARGA", "Valparaíso"),
    ("cartagena-valparaiso", "CARTAGENA", "Valparaíso"),
    ("casablanca-valparaiso", "CASABLANCA", "Valparaíso"),
    ("catemu-valparaiso", "CATEMU", "Valparaíso"),
    ("concon-valparaiso", "CONCON", "Valparaíso"),
    ("el-quisco-valparaiso", "EL QUISCO", "Valparaíso"),
    ("el-tabo-valparaiso", "EL TABO", "Valparaíso"),
    ("hijuelas-valparaiso", "HIJUELAS", "Valparaíso"),
    ("isla-de-pascua-valparaiso", "ISLA DE PASCUA", "Valparaíso"),
    ("juan-fernandez-valparaiso", "JUAN FERNANDEZ", "Valparaíso"),
    ("la-calera-valparaiso", "LA CALERA", "Valparaíso"),
    ("la-cruz-valparaiso", "LA CRUZ", "Valparaíso"),
    ("la-ligua-valparaiso", "LA LIGUA", "Valparaíso"),
    ("limache-valparaiso", "LIMACHE", "Valparaíso"),
    ("llaillay-valparaiso", "LLAILLAY", "Valparaíso"),
    ("los-andes-valparaiso", "LOS ANDES", "Valparaíso"),
    ("nogales-valparaiso", "NOGALES", "Valparaíso"),
    ("olmue-valparaiso", "OLMUE", "Valparaíso"),
    ("panquehue-valparaiso", "PANQUEHUE", "Valparaíso"),
    ("papudo-valparaiso", "PAPUDO", "Valparaíso"),
    ("petorca-valparaiso", "PETORCA", "Valparaíso"),
    ("puchuncavi-valparaiso", "PUCHUNCAVI", "Valparaíso"),
    ("putaendo-valparaiso", "PUTAENDO", "Valparaíso"),
    ("quillota-valparaiso", "QUILLOTA", "Valparaíso"),
    ("quilpue-valparaiso", "QUILPUE", "Valparaíso"),
    ("quintero-valparaiso", "QUINTERO", "Valparaíso"),
    ("rinconada-valparaiso", "RINCONADA", "Valparaíso"),
    ("san-antonio-valparaiso", "SAN ANTONIO", "Valparaíso"),
    ("san-esteban-valparaiso", "SAN ESTEBAN", "Valparaíso"),
    ("san-felipe-valparaiso", "SAN FELIPE", "Valparaíso"),
    ("santa-maria-valparaiso", "SANTA MARIA", "Valparaíso"),
    ("santo-domingo-valparaiso", "SANTO DOMINGO", "Valparaíso"),
    ("valparaiso-valparaiso", "VALPARAISO", "Valparaíso"),
    ("villa-alemana-valparaiso", "VILLA ALEMANA", "Valparaíso"),
    ("vina-del-mar-valparaiso", "VINA DEL MAR", "Valparaíso"),
    ("zapallar-valparaiso", "ZAPALLAR", "Valparaíso"),
    # ---- Metropolitana ----
    ("alhue-metropolitana", "ALHUE", "Metropolitana"),
    ("buin-metropolitana", "BUIN", "Metropolitana"),
    ("calera-de-tango-metropolitana", "CALERA DE TANGO", "Metropolitana"),
    ("cerrillos-metropolitana", "CERRILLOS", "Metropolitana"),
    ("cerro-navia-metropolitana", "CERRO NAVIA", "Metropolitana"),
    ("colina-metropolitana", "COLINA", "Metropolitana"),
    ("conchali-metropolitana", "CONCHALI", "Metropolitana"),
    ("curacavi-metropolitana", "CURACAVI", "Metropolitana"),
    ("el-bosque-metropolitana", "EL BOSQUE", "Metropolitana"),
    ("el-monte-metropolitana", "EL MONTE", "Metropolitana"),
    ("estacion-central-metropolitana", "ESTACION CENTRAL", "Metropolitana"),
    ("huechuraba-metropolitana", "HUECHURABA", "Metropolitana"),
    ("independencia-metropolitana", "INDEPENDENCIA", "Metropolitana"),
    ("isla-de-maipo-metropolitana", "ISLA DE MAIPO", "Metropolitana"),
    ("la-cisterna-metropolitana", "LA CISTERNA", "Metropolitana"),
    ("la-florida-metropolitana", "LA FLORIDA", "Metropolitana"),
    ("la-granja-metropolitana", "LA GRANJA", "Metropolitana"),
    ("la-pintana-metropolitana", "LA PINTANA", "Metropolitana"),
    ("la-reina-metropolitana", "LA REINA", "Metropolitana"),
    ("lampa-metropolitana", "LAMPA", "Metropolitana"),
    ("las-condes-metropolitana", "LAS CONDES", "Metropolitana"),
    ("lo-barnechea-metropolitana", "LO BARNECHEA", "Metropolitana"),
    ("lo-espejo-metropolitana", "LO ESPEJO", "Metropolitana"),
    ("lo-prado-metropolitana", "LO PRADO", "Metropolitana"),
    ("macul-metropolitana", "MACUL", "Metropolitana"),
    ("maipu-metropolitana", "MAIPU", "Metropolitana"),
    ("maria-pinto-metropolitana", "MARIA PINTO", "Metropolitana"),
    ("melipilla-metropolitana", "MELIPILLA", "Metropolitana"),
    ("padre-hurtado-metropolitana", "PADRE HURTADO", "Metropolitana"),
    ("paine-metropolitana", "PAINE", "Metropolitana"),
    ("pedro-aguirre-cerda-metropolitana", "PEDRO AGUIRRE CERDA", "Metropolitana"),
    ("penaflor-metropolitana", "PENAFLOR", "Metropolitana"),
    ("penalolen-metropolitana", "PENALOLEN", "Metropolitana"),
    ("pirque-metropolitana", "PIRQUE", "Metropolitana"),
    ("providencia-metropolitana", "PROVIDENCIA", "Metropolitana"),
    ("pudahuel-metropolitana", "PUDAHUEL", "Metropolitana"),
    ("puente-alto-metropolitana", "PUENTE ALTO", "Metropolitana"),
    ("quilicura-metropolitana", "QUILICURA", "Metropolitana"),
    ("quinta-normal-metropolitana", "QUINTA NORMAL", "Metropolitana"),
    ("recoleta-metropolitana", "RECOLETA", "Metropolitana"),
    ("renca-metropolitana", "RENCA", "Metropolitana"),
    ("san-bernardo-metropolitana", "SAN BERNARDO", "Metropolitana"),
    ("san-joaquin-metropolitana", "SAN JOAQUIN", "Metropolitana"),
    ("san-jose-de-maipo-metropolitana", "SAN JOSE DE MAIPO", "Metropolitana"),
    ("san-miguel-metropolitana", "SAN MIGUEL", "Metropolitana"),
    ("san-pedro-metropolitana", "SAN PEDRO", "Metropolitana"),
    ("san-ramon-metropolitana", "SAN RAMON", "Metropolitana"),
    ("santiago-metropolitana", "SANTIAGO", "Metropolitana"),
    ("talagante-metropolitana", "TALAGANTE", "Metropolitana"),
    ("tiltil-metropolitana", "TILTIL", "Metropolitana"),
    ("vitacura-metropolitana", "VITACURA", "Metropolitana"),
    ("nunoa-metropolitana", "NUNOA", "Metropolitana"),
    # ---- O'Higgins ----
    ("chepica-ohiggins", "CHEPICA", "O'Higgins"),
    ("chimbarongo-ohiggins", "CHIMBARONGO", "O'Higgins"),
    ("codegua-ohiggins", "CODEGUA", "O'Higgins"),
    ("coinco-ohiggins", "COINCO", "O'Higgins"),
    ("coltauco-ohiggins", "COLTAUCO", "O'Higgins"),
    ("donihue-ohiggins", "DONIHUE", "O'Higgins"),
    ("graneros-ohiggins", "GRANEROS", "O'Higgins"),
    ("la-estrella-ohiggins", "LA ESTRELLA", "O'Higgins"),
    ("las-cabras-ohiggins", "LAS CABRAS", "O'Higgins"),
    ("litueche-ohiggins", "LITUECHE", "O'Higgins"),
    ("lolol-ohiggins", "LOLOL", "O'Higgins"),
    ("machali-ohiggins", "MACHALI", "O'Higgins"),
    ("malloa-ohiggins", "MALLOA", "O'Higgins"),
    ("marchihue-ohiggins", "MARCHIHUE", "O'Higgins"),
    ("mostazal-ohiggins", "MOSTAZAL", "O'Higgins"),
    ("nancagua-ohiggins", "NANCAGUA", "O'Higgins"),
    ("navidad-ohiggins", "NAVIDAD", "O'Higgins"),
    ("olivar-ohiggins", "OLIVAR", "O'Higgins"),
    ("palmilla-ohiggins", "PALMILLA", "O'Higgins"),
    ("paredones-ohiggins", "PAREDONES", "O'Higgins"),
    ("peralillo-ohiggins", "PERALILLO", "O'Higgins"),
    ("peumo-ohiggins", "PEUMO", "O'Higgins"),
    ("pichidegua-ohiggins", "PICHIDEGUA", "O'Higgins"),
    ("pichilemu-ohiggins", "PICHILEMU", "O'Higgins"),
    ("placilla-ohiggins", "PLACILLA", "O'Higgins"),
    ("pumanque-ohiggins", "PUMANQUE", "O'Higgins"),
    ("quinta-de-tilcoco-ohiggins", "QUINTA DE TILCOCO", "O'Higgins"),
    ("rancagua-ohiggins", "RANCAGUA", "O'Higgins"),
    ("rengo-ohiggins", "RENGO", "O'Higgins"),
    ("requinoa-ohiggins", "REQUINOA", "O'Higgins"),
    ("san-fernando-ohiggins", "SAN FERNANDO", "O'Higgins"),
    ("san-vicente-ohiggins", "SAN VICENTE", "O'Higgins"),
    ("santa-cruz-ohiggins", "SANTA CRUZ", "O'Higgins"),
    # ---- Maule ----
    ("cauquenes-maule", "CAUQUENES", "Maule"),
    ("chanco-maule", "CHANCO", "Maule"),
    ("colbun-maule", "COLBUN", "Maule"),
    ("constitucion-maule", "CONSTITUCION", "Maule"),
    ("curepto-maule", "CUREPTO", "Maule"),
    ("curico-maule", "CURICO", "Maule"),
    ("empedrado-maule", "EMPEDRADO", "Maule"),
    ("hualane-maule", "HUALANE", "Maule"),
    ("licanten-maule", "LICANTEN", "Maule"),
    ("linares-maule", "LINARES", "Maule"),
    ("longavi-maule", "LONGAVI", "Maule"),
    ("maule-maule", "MAULE", "Maule"),
    ("molina-maule", "MOLINA", "Maule"),
    ("parral-maule", "PARRAL", "Maule"),
    ("pelarco-maule", "PELARCO", "Maule"),
    ("pelluhue-maule", "PELLUHUE", "Maule"),
    ("pencahue-maule", "PENCAHUE", "Maule"),
    ("rauco-maule", "RAUCO", "Maule"),
    ("retiro-maule", "RETIRO", "Maule"),
    ("rio-claro-maule", "RIO CLARO", "Maule"),
    ("romeral-maule", "ROMERAL", "Maule"),
    ("sagrada-familia-maule", "SAGRADA FAMILIA", "Maule"),
    ("san-clemente-maule", "SAN CLEMENTE", "Maule"),
    ("san-javier-maule", "SAN JAVIER", "Maule"),
    ("san-rafael-maule", "SAN RAFAEL", "Maule"),
    ("talca-maule", "TALCA", "Maule"),
    ("teno-maule", "TENO", "Maule"),
    ("vichuquen-maule", "VICHUQUEN", "Maule"),
    ("villa-alegre-maule", "VILLA ALEGRE", "Maule"),
    ("yerbas-buenas-maule", "YERBAS BUENAS", "Maule"),
    # ---- Ñuble ----
    ("bulnes-nuble", "BULNES", "Ñuble"),
    ("chillan-nuble", "CHILLAN", "Ñuble"),
    ("chillan-viejo-nuble", "CHILLAN VIEJO", "Ñuble"),
    ("cobquecura-nuble", "COBQUECURA", "Ñuble"),
    ("coelemu-nuble", "COELEMU", "Ñuble"),
    ("coihueco-nuble", "COIHUECO", "Ñuble"),
    ("el-carmen-nuble", "EL CARMEN", "Ñuble"),
    ("ninhue-nuble", "NINHUE", "Ñuble"),
    ("niquen-nuble", "NIQUEN", "Ñuble"),
    ("pemuco-nuble", "PEMUCO", "Ñuble"),
    ("pinto-nuble", "PINTO", "Ñuble"),
    ("portezuelo-nuble", "PORTEZUELO", "Ñuble"),
    ("quillon-nuble", "QUILLON", "Ñuble"),
    ("quirihue-nuble", "QUIRIHUE", "Ñuble"),
    ("ranquil-nuble", "RANQUIL", "Ñuble"),
    ("san-carlos-nuble", "SAN CARLOS", "Ñuble"),
    ("san-fabian-nuble", "SAN FABIAN", "Ñuble"),
    ("san-ignacio-nuble", "SAN IGNACIO", "Ñuble"),
    ("san-nicolas-nuble", "SAN NICOLAS", "Ñuble"),
    ("treguaco-nuble", "TREGUACO", "Ñuble"),
    ("yungay-nuble", "YUNGAY", "Ñuble"),
    # ---- Biobío ----
    ("alto-biobio-biobio", "ALTO BIOBIO", "Biobío"),
    ("antuco-biobio", "ANTUCO", "Biobío"),
    ("arauco-biobio", "ARAUCO", "Biobío"),
    ("cabrero-biobio", "CABRERO", "Biobío"),
    ("canete-biobio", "CANETE", "Biobío"),
    ("chiguayante-biobio", "CHIGUAYANTE", "Biobío"),
    ("concepcion-biobio", "CONCEPCION", "Biobío"),
    ("contulmo-biobio", "CONTULMO", "Biobío"),
    ("coronel-biobio", "CORONEL", "Biobío"),
    ("curanilahue-biobio", "CURANILAHUE", "Biobío"),
    ("florida-biobio", "FLORIDA", "Biobío"),
    ("hualpen-biobio", "HUALPEN", "Biobío"),
    ("hualqui-biobio", "HUALQUI", "Biobío"),
    ("laja-biobio", "LAJA", "Biobío"),
    ("lebu-biobio", "LEBU", "Biobío"),
    ("los-angeles-biobio", "LOS ANGELES", "Biobío"),
    ("los-alamos-biobio", "LOS ALAMOS", "Biobío"),
    ("lota-biobio", "LOTA", "Biobío"),
    ("mulchen-biobio", "MULCHEN", "Biobío"),
    ("nacimiento-biobio", "NACIMIENTO", "Biobío"),
    ("negrete-biobio", "NEGRETE", "Biobío"),
    ("penco-biobio", "PENCO", "Biobío"),
    ("quilaco-biobio", "QUILACO", "Biobío"),
    ("quilleco-biobio", "QUILLECO", "Biobío"),
    ("san-pedro-de-la-paz-biobio", "SAN PEDRO DE LA PAZ", "Biobío"),
    ("san-rosendo-biobio", "SAN ROSENDO", "Biobío"),
    ("santa-barbara-biobio", "SANTA BARBARA", "Biobío"),
    ("santa-juana-biobio", "SANTA JUANA", "Biobío"),
    ("talcahuano-biobio", "TALCAHUANO", "Biobío"),
    ("tirua-biobio", "TIRUA", "Biobío"),
    ("tome-biobio", "TOME", "Biobío"),
    ("tucapel-biobio", "TUCAPEL", "Biobío"),
    ("yumbel-biobio", "YUMBEL", "Biobío"),
    # ---- Araucanía ----
    ("angol-araucania", "ANGOL", "Araucanía"),
    ("carahue-araucania", "CARAHUE", "Araucanía"),
    ("cholchol-araucania", "CHOLCHOL", "Araucanía"),
    ("collipulli-araucania", "COLLIPULLI", "Araucanía"),
    ("cunco-araucania", "CUNCO", "Araucanía"),
    ("curacautin-araucania", "CURACAUTIN", "Araucanía"),
    ("curarrehue-araucania", "CURARREHUE", "Araucanía"),
    ("ercilla-araucania", "ERCILLA", "Araucanía"),
    ("freire-araucania", "FREIRE", "Araucanía"),
    ("galvarino-araucania", "GALVARINO", "Araucanía"),
    ("gorbea-araucania", "GORBEA", "Araucanía"),
    ("lautaro-araucania", "LAUTARO", "Araucanía"),
    ("loncoche-araucania", "LONCOCHE", "Araucanía"),
    ("lonquimay-araucania", "LONQUIMAY", "Araucanía"),
    ("los-sauces-araucania", "LOS SAUCES", "Araucanía"),
    ("lumaco-araucania", "LUMACO", "Araucanía"),
    ("melipeuco-araucania", "MELIPEUCO", "Araucanía"),
    ("nueva-imperial-araucania", "NUEVA IMPERIAL", "Araucanía"),
    ("padre-las-casas-araucania", "PADRE LAS CASAS", "Araucanía"),
    ("perquenco-araucania", "PERQUENCO", "Araucanía"),
    ("pitrufquen-araucania", "PITRUFQUEN", "Araucanía"),
    ("pucon-araucania", "PUCON", "Araucanía"),
    ("puren-araucania", "PUREN", "Araucanía"),
    ("renaico-araucania", "RENAICO", "Araucanía"),
    ("saavedra-araucania", "SAAVEDRA", "Araucanía"),
    ("teodoro-schmidt-araucania", "TEODORO SCHMIDT", "Araucanía"),
    ("temuco-araucania", "TEMUCO", "Araucanía"),
    ("tolten-araucania", "TOLTEN", "Araucanía"),
    ("traiguen-araucania", "TRAIGUEN", "Araucanía"),
    ("victoria-araucania", "VICTORIA", "Araucanía"),
    ("vilcun-araucania", "VILCUN", "Araucanía"),
    ("villarrica-araucania", "VILLARRICA", "Araucanía"),
    # ---- Los Ríos ----
    ("corral-los-rios", "CORRAL", "Los Ríos"),
    ("futrono-los-rios", "FUTRONO", "Los Ríos"),
    ("la-union-los-rios", "LA UNION", "Los Ríos"),
    ("lago-ranco-los-rios", "LAGO RANCO", "Los Ríos"),
    ("lanco-los-rios", "LANCO", "Los Ríos"),
    ("los-lagos-los-rios", "LOS LAGOS", "Los Ríos"),
    ("mafil-los-rios", "MAFIL", "Los Ríos"),
    ("mariquina-los-rios", "MARIQUINA", "Los Ríos"),
    ("paillaco-los-rios", "PAILLACO", "Los Ríos"),
    ("panguipulli-los-rios", "PANGUIPULLI", "Los Ríos"),
    ("rio-bueno-los-rios", "RIO BUENO", "Los Ríos"),
    ("valdivia-los-rios", "VALDIVIA", "Los Ríos"),
    # ---- Los Lagos ----
    ("ancud-los-lagos", "ANCUD", "Los Lagos"),
    ("calbuco-los-lagos", "CALBUCO", "Los Lagos"),
    ("castro-los-lagos", "CASTRO", "Los Lagos"),
    ("chaiten-los-lagos", "CHAITEN", "Los Lagos"),
    ("chonchi-los-lagos", "CHONCHI", "Los Lagos"),
    ("cochamo-los-lagos", "COCHAMO", "Los Lagos"),
    ("curaco-de-velez-los-lagos", "CURACO DE VELEZ", "Los Lagos"),
    ("dalcahue-los-lagos", "DALCAHUE", "Los Lagos"),
    ("fresia-los-lagos", "FRESIA", "Los Lagos"),
    ("frutillar-los-lagos", "FRUTILLAR", "Los Lagos"),
    ("futaleufu-los-lagos", "FUTALEUFU", "Los Lagos"),
    ("hualaihue-los-lagos", "HUALAIHUE", "Los Lagos"),
    ("llanquihue-los-lagos", "LLANQUIHUE", "Los Lagos"),
    ("los-muermos-los-lagos", "LOS MUERMOS", "Los Lagos"),
    ("maullin-los-lagos", "MAULLIN", "Los Lagos"),
    ("osorno-los-lagos", "OSORNO", "Los Lagos"),
    ("palena-los-lagos", "PALENA", "Los Lagos"),
    ("puerto-montt-los-lagos", "PUERTO MONTT", "Los Lagos"),
    ("puerto-octay-los-lagos", "PUERTO OCTAY", "Los Lagos"),
    ("puerto-varas-los-lagos", "PUERTO VARAS", "Los Lagos"),
    ("puqueldon-los-lagos", "PUQUELDON", "Los Lagos"),
    ("purranque-los-lagos", "PURRANQUE", "Los Lagos"),
    ("puyehue-los-lagos", "PUYEHUE", "Los Lagos"),
    ("queilen-los-lagos", "QUEILEN", "Los Lagos"),
    ("quellon-los-lagos", "QUELLON", "Los Lagos"),
    ("quemchi-los-lagos", "QUEMCHI", "Los Lagos"),
    ("quinchao-los-lagos", "QUINCHAO", "Los Lagos"),
    ("rio-negro-los-lagos", "RIO NEGRO", "Los Lagos"),
    ("san-juan-de-la-costa-los-lagos", "SAN JUAN DE LA COSTA", "Los Lagos"),
    ("san-pablo-los-lagos", "SAN PABLO", "Los Lagos"),
    # ---- Aysén ----
    ("aysen-aysen", "AYSEN", "Aysén"),
    ("chile-chico-aysen", "CHILE CHICO", "Aysén"),
    ("cisnes-aysen", "CISNES", "Aysén"),
    ("cochrane-aysen", "COCHRANE", "Aysén"),
    ("coyhaique-aysen", "COYHAIQUE", "Aysén"),
    ("guaitecas-aysen", "GUAITECAS", "Aysén"),
    ("lago-verde-aysen", "LAGO VERDE", "Aysén"),
    ("ohiggins-aysen", "OHIGGINS", "Aysén"),
    ("rio-ibanez-aysen", "RIO IBANEZ", "Aysén"),
    ("tortel-aysen", "TORTEL", "Aysén"),
    # ---- Magallanes ----
    ("antartica-magallanes", "ANTARTICA", "Magallanes"),
    ("cabo-de-hornos-magallanes", "CABO DE HORNOS", "Magallanes"),
    ("laguna-blanca-magallanes", "LAGUNA BLANCA", "Magallanes"),
    ("natales-magallanes", "NATALES", "Magallanes"),
    ("porvenir-magallanes", "PORVENIR", "Magallanes"),
    ("primavera-magallanes", "PRIMAVERA", "Magallanes"),
    ("punta-arenas-magallanes", "PUNTA ARENAS", "Magallanes"),
    ("rio-verde-magallanes", "RIO VERDE", "Magallanes"),
    ("san-gregorio-magallanes", "SAN GREGORIO", "Magallanes"),
    ("timaukel-magallanes", "TIMAUKEL", "Magallanes"),
    ("torres-del-paine-magallanes", "TORRES DEL PAINE", "Magallanes"),
]

TIPOS = ["casa", "departamento"]


def main():
    print("=" * 70)
    print(" BARRIDO NACIONAL DE PORTAL INMOBILIARIO - LAS 346 COMUNAS DE CHILE")
    print("=" * 70)

    pendientes = [(slug, nombre, region) for slug, nombre, region in TODAS_LAS_COMUNAS_CHILE
                  if slug not in _YA_CUBIERTAS]
    print(f"Comunas ya cubiertas por el barrido diario (se saltan): {len(_YA_CUBIERTAS)}")
    print(f"Comunas nuevas a recorrer ahora: {len(pendientes)}")
    print(f"Total de páginas a pedir: {len(pendientes) * len(TIPOS)} "
          f"(a ~{PAUSA_SEG}s de pausa cada una, calcula el tiempo total)")
    print()

    conexion, cursor = pm.conectar_bd_mercado(CARPETA_DATOS)
    hoy = dt.date.today().isoformat()

    resumen_por_region = {}  # región -> [comunas_con_datos, comunas_en_cero]
    nuevas_totales = ya_vistas_totales = descartadas_totales = 0

    for i, (slug, nombre, region) in enumerate(pendientes, start=1):
        encontro_algo = False
        for tipo in TIPOS:
            url = f"https://www.portalinmobiliario.com/venta/{tipo}/{slug}"
            try:
                r = requests.get(url, headers=pm.HEADERS, timeout=20)
                r.raise_for_status()
                publicaciones = pm._tarjetas_genericas(r.text, url, comuna_fija=nombre)
            except Exception as e:
                publicaciones = []
                print(f"  [{i}/{len(pendientes)}] {nombre} ({tipo}) -> FALLÓ: {type(e).__name__}: {e}")
                time.sleep(PAUSA_SEG)
                continue

            nuevas = ya_vistas = descartadas = 0
            for pub in publicaciones:
                resultado = pm.guardar_publicacion(
                    cursor, "PortalInmobiliario", pub["comuna"], pub["tipo_propiedad"],
                    pub["precio_m2"], pub["año_publicacion"], pub.get("url"), hoy,
                )
                if resultado == "nueva":
                    nuevas += 1
                elif resultado == "ya_existia":
                    ya_vistas += 1
                else:
                    descartadas += 1
            conexion.commit()
            nuevas_totales += nuevas
            ya_vistas_totales += ya_vistas
            descartadas_totales += descartadas
            if nuevas + ya_vistas > 0:
                encontro_algo = True

            print(f"  [{i}/{len(pendientes)}] {nombre} ({tipo}): "
                  f"{len(publicaciones)} publicaciones -> {nuevas} nuevas, {ya_vistas} ya vistas")
            time.sleep(PAUSA_SEG)

        resumen_por_region.setdefault(region, [0, 0])
        if encontro_algo:
            resumen_por_region[region][0] += 1
        else:
            resumen_por_region[region][1] += 1

    conexion.close()

    print()
    print("=" * 70)
    print(" RESUMEN FINAL")
    print("=" * 70)
    print(f"Total nuevas publicaciones guardadas: {nuevas_totales}")
    print(f"Total ya vistas (de una corrida anterior): {ya_vistas_totales}")
    print(f"Total descartadas (datos incompletos): {descartadas_totales}")
    print()
    print("Por región (comunas con al menos una publicación encontrada / total probadas):")
    for region, (con_datos, en_cero) in sorted(resumen_por_region.items()):
        total = con_datos + en_cero
        print(f"  {region}: {con_datos}/{total} comunas con datos"
              + (f"  <-- revisar, {en_cero} en cero" if en_cero > 0 else ""))
    print()
    print("Listo. Mándame este resumen completo (no hace falta pegar las 690 líneas,")
    print("con el resumen final por región basta) para que ajustemos lo que haya fallado.")


if __name__ == "__main__":
    main()
