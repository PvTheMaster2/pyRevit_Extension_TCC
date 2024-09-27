# -*- coding: utf-8 -*-
__title__ = "Inserir Tomada com Parâmetros Elétricos"
__doc__ = """Versão: 2.0
_____________________________________________________________________
Descrição:
Este script insere uma tomada elétrica na parede selecionada,
permitindo escolher a altura, posição horizontal e face (frontal/traseira).
Adicionalmente, coleta parâmetros elétricos como potência aparente,
fator de potência, tensão e número de fases para calcular a potência ativa.
_____________________________________________________________________
Como usar:
- Clique no botão e siga as instruções.
_____________________________________________________________________
Autor: Seu Nome"""

# Importações necessárias
import clr
clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import DialogResult, MessageBox, MessageBoxButtons, MessageBoxIcon

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.Exceptions import InvalidOperationException
from Autodesk.Revit.DB.Structure import StructuralType

# Importações do pyRevit
from pyrevit import revit, forms, script

# Variáveis do documento
doc = __revit__.ActiveUIDocument.Document  # type: Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

# Funções auxiliares
def selecionar_familia_tomada():
    """Permite que o usuário selecione uma família de tomada elétrica."""
    # Coletar todos os símbolos de família da categoria "Dispositivos elétricos"
    collector = FilteredElementCollector(doc) \
        .OfClass(FamilySymbol) \
        .OfCategory(BuiltInCategory.OST_ElectricalFixtures)

    tomadas = []
    for symbol in collector:
        # Obter o nome da família usando BuiltInParameter
        family_name_param = symbol.get_Parameter(BuiltInParameter.ALL_MODEL_FAMILY_NAME)
        if family_name_param and family_name_param.HasValue:
            family_name = family_name_param.AsString()
        else:
            family_name = "Sem Família"

        # Obter o nome do símbolo usando BuiltInParameter
        symbol_name_param = symbol.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
        if symbol_name_param and symbol_name_param.HasValue:
            symbol_name = symbol_name_param.AsString()
        else:
            symbol_name = "Sem Nome"

        # Filtrar por famílias que contenham "Tomada" no nome
        if "Tomada" in family_name or "Tomada" in symbol_name:
            tomadas.append(symbol)

    if not tomadas:
        forms.alert("Nenhuma família de tomadas encontrada no projeto.", exitscript=True)

    # Criar um dicionário de opções
    tomadas_dict = {}
    for tomada in tomadas:
        try:
            # Obter o nome da família
            family_name_param = tomada.get_Parameter(BuiltInParameter.ALL_MODEL_FAMILY_NAME)
            if family_name_param and family_name_param.HasValue:
                family_name = family_name_param.AsString()
            else:
                family_name = "Sem Família"

            # Obter o nome do símbolo
            symbol_name_param = tomada.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
            if symbol_name_param and symbol_name_param.HasValue:
                symbol_name = symbol_name_param.AsString()
            else:
                symbol_name = "Sem Nome"

            display_name = "{} : {}".format(family_name, symbol_name)
            tomadas_dict[display_name] = tomada
        except Exception as e:
            # Se ocorrer um erro, ignorar este elemento
            pass

    if not tomadas_dict:
        forms.alert("Nenhuma família de tomadas válida encontrada.", exitscript=True)

    # Ordenar os nomes para exibição
    tomadas_nomes_ordenados = sorted(tomadas_dict.keys())

    # Permitir que o usuário selecione uma tomada
    tomada_selecionada_nome = forms.SelectFromList.show(
        tomadas_nomes_ordenados,
        title='Selecione uma Tomada',
        button_name='Selecionar',
        multiselect=False
    )

    if not tomada_selecionada_nome:
        forms.alert("Nenhuma tomada selecionada.", exitscript=True)

    tomada_selecionada = tomadas_dict[tomada_selecionada_nome]

    # Ativar o símbolo da família, se necessário
    if not tomada_selecionada.IsActive:
        with revit.Transaction("Ativar Família"):
            tomada_selecionada.Activate()
            doc.Regenerate()

    return tomada_selecionada

def selecionar_parede():
    """Permite que o usuário selecione uma parede."""
    sel = uidoc.Selection
    try:
        referencia = sel.PickObject(ObjectType.Element, 'Selecione a parede onde a tomada será inserida.')
        parede = doc.GetElement(referencia.ElementId)
    except InvalidOperationException:
        forms.alert("Nenhuma parede selecionada.", exitscript=True)

    if not isinstance(parede, Wall):
        forms.alert("O elemento selecionado não é uma parede.", exitscript=True)

    return parede

def obter_parametros_elec():
    """Obtém os parâmetros elétricos do usuário."""
    # Obter a potência aparente (S)
    potencia_aparente_input = forms.ask_for_string(
        prompt="Insira a potência aparente (S) em VA ou kVA:",
        title="Potência Aparente",
        default="1000"  # Valor padrão em VA
    )
    try:
        potencia_aparente = float(potencia_aparente_input.replace(',', '.'))
    except ValueError:
        forms.alert("Entrada inválida. Usando potência aparente padrão de 1000 VA.")
        potencia_aparente = 1000.0  # Valor padrão

    # Obter o fator de potência (cos φ)
    fator_potencia_input = forms.ask_for_string(
        prompt="Insira o fator de potência (cos φ):",
        title="Fator de Potência",
        default="0.8"
    )
    try:
        fator_potencia = float(fator_potencia_input.replace(',', '.'))
        if not (0 < fator_potencia <= 1):
            raise ValueError
    except ValueError:
        forms.alert("Entrada inválida. Usando fator de potência padrão de 0.8.")
        fator_potencia = 0.8

    # Obter a tensão (V)
    tensao_input = forms.ask_for_string(
        prompt="Insira a tensão (V):",
        title="Tensão",
        default="127"
    )
    try:
        tensao = float(tensao_input.replace(',', '.'))
    except ValueError:
        forms.alert("Entrada inválida. Usando tensão padrão de 127 V.")
        tensao = 127.0

    # Obter o número de fases
    numero_fases_input = forms.ask_for_string(
        prompt="Insira o número de fases (1 para monofásico, 3 para trifásico):",
        title="Número de Fases",
        default="1"
    )
    try:
        numero_fases = int(numero_fases_input)
        if numero_fases not in [1, 3]:
            raise ValueError
    except ValueError:
        forms.alert("Entrada inválida. Usando número de fases padrão de 1 (monofásico).")
        numero_fases = 1

    # Calcular a potência ativa (P)
    potencia_ativa = potencia_aparente * fator_potencia  # P = S * cos φ

    # Opcional: Calcular a potência reativa (Q) se desejar
    # potencia_reativa = potencia_aparente * (1 - fator_potencia**2)**0.5  # Q = S * sin φ

    # Mostrar os cálculos ao usuário (substituir f-string por .format())
    MessageBox.Show(
        "Potência Ativa (P): {} W".format(potencia_ativa),
        "Cálculo da Potência Ativa",
        MessageBoxButtons.OK,
        MessageBoxIcon.Information
    )

    return potencia_aparente, fator_potencia, tensao, numero_fases, potencia_ativa

def obter_ponto_insercao(parede):
    """Obtém o ponto de inserção da tomada na parede."""
    # Obter a altura desejada do usuário
    altura_metros_input = forms.ask_for_string(
        prompt="Insira a altura da tomada em metros:",
        title="Altura da Tomada",
        default="1.10"
    )
    try:
        altura_metros = float(altura_metros_input.replace(',', '.'))
    except ValueError:
        forms.alert("Entrada inválida. Usando altura padrão de 1.10 metros.")
        altura_metros = 1.10

    # Converter metros para pés
    altura_pes = altura_metros * 3.28084

    # Obter a localização da parede
    loc_curve = parede.Location
    if not isinstance(loc_curve, LocationCurve):
        forms.alert("Não foi possível obter a localização da parede.", exitscript=True)

    # Obter a curva (linha) central da parede
    curva = loc_curve.Curve

    # Solicitar ao usuário a distância ao longo da parede
    distancia_metros_input = forms.ask_for_string(
        prompt="Insira a distância ao longo da parede em metros (0 para início, deixe em branco para meio):",
        title="Posição Horizontal",
        default=""
    )
    if distancia_metros_input.strip() == "":
        # Usar o ponto médio se o usuário não inserir nada
        parametro_normalizado = 0.5
    else:
        try:
            distancia_metros = float(distancia_metros_input.replace(',', '.'))
            comprimento_parede = curva.Length  # Comprimento em pés
            comprimento_metros = comprimento_parede / 3.28084
            parametro_normalizado = distancia_metros / comprimento_metros
            parametro_normalizado = max(0.0, min(1.0, parametro_normalizado))  # Garantir que esteja entre 0 e 1
        except ValueError:
            forms.alert("Entrada inválida. Usando o ponto médio da parede.")
            parametro_normalizado = 0.5

    # Obter o ponto ao longo da curva
    ponto_na_curva = curva.Evaluate(parametro_normalizado, True)

    # Ajustar o ponto para a altura desejada
    ponto_insercao = XYZ(ponto_na_curva.X, ponto_na_curva.Y, ponto_na_curva.Z + altura_pes)

    # Obter a espessura da parede
    espessura_parede = parede.WallType.Width  # Em pés

    # Perguntar ao usuário em qual face deseja inserir a tomada
    face_opcoes = ['Frontal', 'Traseira']
    face_selecionada = forms.SelectFromList.show(
        face_opcoes,
        title='Selecione a Face da Parede',
        button_name='Selecionar',
        multiselect=False
    )

    if not face_selecionada:
        forms.alert("Nenhuma face selecionada. Usando a posição central da parede.")
        face_selecionada = None  # Indica que não foi selecionada

    else:
        # Calcular o vetor normal da parede
        direcao_parede = (curva.GetEndPoint(1) - curva.GetEndPoint(0)).Normalize()
        vetor_normal = XYZ(-direcao_parede.Y, direcao_parede.X, 0).Normalize()

        # Calcular o deslocamento (metade da espessura da parede)
        deslocamento = (espessura_parede / 2)

        if face_selecionada == 'Frontal':
            # Deslocar na direção do vetor normal
            ponto_insercao += vetor_normal * deslocamento
        elif face_selecionada == 'Traseira':
            # Deslocar na direção oposta ao vetor normal
            ponto_insercao -= vetor_normal * deslocamento

    return ponto_insercao

def inserir_tomada(parede, tomada_selecionada, ponto_insercao, parametros_elec):
    """Insere a tomada nas posições calculadas e armazena os parâmetros elétricos."""
    potencia_aparente, fator_potencia, tensao, numero_fases, potencia_ativa = parametros_elec

    # Inserir a tomada usando a parede como host
    with revit.Transaction("Inserir Tomada"):
        tomada_instancia = doc.Create.NewFamilyInstance(
            ponto_insercao,
            tomada_selecionada,
            parede,
            StructuralType.NonStructural
        )

        # Ajustar a orientação da tomada para ficar alinhada com a parede
        # Obter a direção da parede
        loc_curve = parede.Location
        curva = loc_curve.Curve
        direcao_parede = (curva.GetEndPoint(1) - curva.GetEndPoint(0)).Normalize()

        # Calcular o ângulo entre a direção da parede e o eixo X
        angulo = XYZ.BasisX.AngleTo(direcao_parede)

        # Determinar o sentido da rotação
        cross = XYZ.BasisX.CrossProduct(direcao_parede)
        if cross.Z < 0:
            angulo = -angulo

        # Criar o eixo de rotação (linha vertical passando pelo ponto de inserção)
        eixo_rotacao = Line.CreateBound(ponto_insercao, ponto_insercao + XYZ.BasisZ)

        # Rotacionar a tomada
        ElementTransformUtils.RotateElement(
            doc,
            tomada_instancia.Id,
            eixo_rotacao,
            angulo
        )

        # **Armazenar os parâmetros elétricos na instância da família (se aplicável)**
        # Verifique se a família possui os parâmetros correspondentes antes de tentar defini-los
        try:
            # Exemplo: supondo que a família tenha parâmetros compartilhados chamados "Potencia_Aparente",
            # "Fator_Potencia", "Tensao", "Numero_Fases", "Potencia_Ativa"

            parametro_S = tomada_instancia.LookupParameter("Potencia_Aparente")
            if parametro_S and parametro_S.StorageType == StorageType.Double:
                parametro_S.Set(potencia_aparente)

            parametro_cos_phi = tomada_instancia.LookupParameter("Fator_Potencia")
            if parametro_cos_phi and parametro_cos_phi.StorageType == StorageType.Double:
                parametro_cos_phi.Set(fator_potencia)

            parametro_V = tomada_instancia.LookupParameter("Tensao")
            if parametro_V and parametro_V.StorageType == StorageType.Double:
                parametro_V.Set(tensao)

            parametro_fases = tomada_instancia.LookupParameter("Numero_Fases")
            if parametro_fases and parametro_fases.StorageType == StorageType.Integer:
                parametro_fases.Set(numero_fases)

            parametro_P = tomada_instancia.LookupParameter("Potencia_Ativa")
            if parametro_P and parametro_P.StorageType == StorageType.Double:
                parametro_P.Set(potencia_ativa)
        except Exception as e:
            # Se ocorrer um erro ao definir os parâmetros, exibir uma mensagem e continuar
            forms.alert("Erro ao definir parâmetros elétricos: {}".format(e))

    forms.alert("Tomada inserida com sucesso!")

def inserir_tomada_na_parede():
    """Função principal para inserir uma tomada na parede com coleta de parâmetros elétricos."""
    try:
        # Selecionar a família de tomada
        tomada_selecionada = selecionar_familia_tomada()

        # Selecionar a parede
        parede = selecionar_parede()

        # Obter os parâmetros elétricos do usuário
        parametros_elec = obter_parametros_elec()

        # Obter o ponto de inserção
        ponto_insercao = obter_ponto_insercao(parede)

        # Inserir a tomada com os parâmetros elétricos
        inserir_tomada(parede, tomada_selecionada, ponto_insercao, parametros_elec)

    except Exception as e:
        forms.alert("Ocorreu um erro: {}".format(e))

# Executar o script
if __name__ == "__main__":
    inserir_tomada_na_parede()
