# -*- coding: utf-8 -*-
__title__ = "Log de Informações da Tomada"
__doc__ = """Script para selecionar uma família de tomada e logar todas as suas informações e parâmetros."""

# Importações necessárias
import clr
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.Exceptions import InvalidOperationException

# Importações do pyRevit
from pyrevit import revit, forms, script

import traceback  # Para capturar o traceback completo

# Variáveis do documento
doc = revit.doc  # Documento ativo do Revit
uidoc = revit.uidoc  # Documento UI ativo


def selecionar_familia_tomada():
    """Permite que o usuário selecione uma família de tomada elétrica."""
    output = script.get_output()
    output.print_md("### Iniciando seleção da família de tomada.")

    # Coletar todos os símbolos de família da categoria "Dispositivos elétricos"
    collector = FilteredElementCollector(doc) \
        .OfClass(FamilySymbol) \
        .OfCategory(BuiltInCategory.OST_ElectricalFixtures)

    tomadas = []
    for symbol in collector:
        try:
            # Obter o nome da família usando BuiltInParameter
            family_name_param = symbol.get_Parameter(BuiltInParameter.ALL_MODEL_FAMILY_NAME)
            family_name = family_name_param.AsString() if family_name_param and family_name_param.HasValue else "Sem Família"

            # Obter o nome do símbolo usando BuiltInParameter
            symbol_name_param = symbol.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
            symbol_name = symbol_name_param.AsString() if symbol_name_param and symbol_name_param.HasValue else "Sem Nome"

            # Filtrar por famílias que contenham "Tomada" no nome
            if "Tomada" in family_name or "Tomada" in symbol_name:
                tomadas.append(symbol)
        except Exception as e:
            output.print_md("**Erro ao processar símbolo de família:** {}".format(e))
            pass

    if not tomadas:
        forms.alert("Nenhuma família de tomadas encontrada no projeto.", exitscript=True)

    # Criar um dicionário de opções
    tomadas_dict = {}
    for tomada in tomadas:
        try:
            # Obter o nome da família
            family_name_param = tomada.get_Parameter(BuiltInParameter.ALL_MODEL_FAMILY_NAME)
            family_name = family_name_param.AsString() if family_name_param and family_name_param.HasValue else "Sem Família"

            # Obter o nome do símbolo
            symbol_name_param = tomada.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
            symbol_name = symbol_name_param.AsString() if symbol_name_param and symbol_name_param.HasValue else "Sem Nome"

            display_name = "{} : {}".format(family_name, symbol_name)
            tomadas_dict[display_name] = tomada
        except Exception as e:
            output.print_md("**Erro ao processar tomada:** {}".format(e))
            pass

    if not tomadas_dict:
        forms.alert("Nenhuma família de tomadas válida encontrada.", exitscript=True)

    # Ordenar os nomes para exibição
    tomadas_nomes_ordenados = sorted(tomadas_dict.keys())

    # Permitir que o usuário selecione uma tomada
    output.print_md("### Exibindo lista de tomadas para seleção.")
    tomada_selecionada_nome = forms.SelectFromList.show(
        tomadas_nomes_ordenados,
        title='Selecione uma Tomada',
        button_name='Selecionar',
        multiselect=False
    )

    if not tomada_selecionada_nome:
        forms.alert("Nenhuma tomada selecionada.", exitscript=True)

    tomada_selecionada = tomadas_dict[tomada_selecionada_nome]

    # Adicionar logs para depuração
    output.print_md("### Tipo de tomada_selecionada: {}".format(type(tomada_selecionada)))
    output.print_md("### Atributos disponíveis: {}".format(dir(tomada_selecionada)))

    # Ativar o símbolo da família, se necessário
    if not tomada_selecionada.IsActive:
        output.print_md("### Ativando o símbolo da família da tomada.")
        with revit.Transaction("Ativar Família"):
            tomada_selecionada.Activate()
            doc.Regenerate()

    # Verificar se o atributo 'Name' existe
    if hasattr(tomada_selecionada, 'Name'):
        output.print_md("### Família de tomada selecionada: {}".format(tomada_selecionada.Name))
    else:
        output.print_md("### O objeto tomada_selecionada não possui o atributo 'Name'.")

    return tomada_selecionada


def logar_informacoes_tomada(tomada):
    """Loga todas as informações e parâmetros da tomada selecionada."""
    output = script.get_output()
    output.print_md("### Iniciando log de informações da tomada.")

    # Logar informações básicas da família
    output.print_md("**Informações da Família:**")
    try:
        family = tomada.Family
        output.print_md("- **Nome da Família:** {}".format(family.Name))
        output.print_md("- **Nome do Tipo:** {}".format(tomada.Name))
        output.print_md("- **IsActive:** {}".format(tomada.IsActive))
    except Exception as e:
        output.print_md("**Erro ao obter informações básicas da família:** {}".format(e))

    # Logar todos os parâmetros da família
    output.print_md("**Parâmetros da Tomada:**")
    try:
        parameters = tomada.Parameters
        for param in parameters:
            param_name = param.Definition.Name
            param_value = ""
            if param.StorageType == StorageType.Double:
                # Converter de pés para metros, se aplicável
                param_value = param.AsDouble()
                # Dependendo do parâmetro, você pode precisar converter unidades
            elif param.StorageType == StorageType.Integer:
                param_value = param.AsInteger()
            elif param.StorageType == StorageType.String:
                param_value = param.AsString()
            elif param.StorageType == StorageType.ElementId:
                param_value = param.AsElementId().IntegerValue
            else:
                param_value = "Tipo de armazenamento desconhecido"

            output.print_md("- **{}:** {}".format(param_name, param_value))
    except Exception as e:
        output.print_md("**Erro ao obter parâmetros da tomada:** {}".format(e))

    # Logar conectores elétricos, se houver
    output.print_md("**Conectores Elétricos da Tomada:**")
    try:
        connectors = tomada.MEPModel.ConnectorManager.Connectors
        if connectors.Size == 0:
            output.print_md("- Nenhum conector encontrado.")
        else:
            for connector in connectors:
                output.print_md("- **Nome do Conector:** {}".format(connector.Name))
                output.print_md("  - **Domain:** {}".format(connector.Domain))
                output.print_md("  - **ConnectorType:** {}".format(connector.ConnectorType))
                output.print_md(
                    "  - **Origin:** ({:.2f}, {:.2f}, {:.2f})".format(connector.Origin.X, connector.Origin.Y,
                                                                      connector.Origin.Z))
                output.print_md("  - **Facing Orientation:** {}".format(connector.FacingOrientation))
                output.print_md("  - **Is Connected:** {}".format(connector.IsConnected))
    except Exception as e:
        output.print_md("**Erro ao obter conectores elétricos:** {}".format(e))


def main():
    try:
        # Selecionar a família de tomada
        tomada_selecionada = selecionar_familia_tomada()

        # Logar informações da tomada
        logar_informacoes_tomada(tomada_selecionada)

        forms.alert("Log de informações da tomada concluído. Verifique o painel de Output do pyRevit.")
    except Exception as e:
        tb = traceback.format_exc()
        forms.alert("Ocorreu um erro:\n{}".format(tb))
        script.get_output().print_md("### Erro Geral no Script:\n{}".format(tb))


# Executar o script
if __name__ == "__main__":
    main()
