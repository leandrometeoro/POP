import sys
from lxml import etree
import json

def parse_bpmn_pop(file_path):
    """
    Analisa um arquivo BPMN do Camunda 8 e extrai os metadados do template POP,
    bem como a documentação das tarefas.

    Args:
        file_path (str): O caminho para o arquivo .bpmn ou .xml.

    Returns:
        dict: Um dicionário contendo os metadados extraídos, ou None se ocorrer um erro.
    """
    print(f"INFO: Analisando o arquivo: {file_path}")

    try:
        # Namespaces corretos para o seu arquivo BPMN
        ns = {
            'bpmn': 'http://www.omg.org/spec/BPMN/20100524/MODEL',
            'zeebe': 'http://camunda.org/schema/zeebe/1.0'
        }

        tree = etree.parse(file_path)
        root = tree.getroot()

        # XPath corrigido para encontrar o participante correto (com os dados preenchidos)
        # e usando o método .xpath() que é mais poderoso
        participant_xpath = f"//bpmn:participant[@name='Registro de Software']"
        participants = root.xpath(participant_xpath, namespaces=ns)

        if not participants:
            print("ERRO: Não foi possível encontrar o participante 'Registro de Software' no arquivo.")
            return None
        
        participant = participants[0] # Pega o primeiro resultado da busca

        # Lista de campos que devem ser tratados como listas separadas por "//"
        multi_value_fields = [
            'palavrasChaveAdicionais',
            'dicionarioAdicionais_termos',
            'dicionarioAdicionais_significados',
            'referenciasAdicionais_refs',
            'referenciasAdicionais_descs',
            'sistemasAdicionais',
            'indicadoresMonAdicionais_nomes',
            'indicadoresMonAdicionais_descs',
            'observacoesAdicionais',
            'riscoDigitado3_adicionais',
            'alteracao_itens'
        ]

        pop_properties = {}
        properties_xpath = ".//zeebe:properties/zeebe:property"
        
        print("INFO: Extraindo propriedades do template POP...")
        for prop in participant.findall(properties_xpath, namespaces=ns):
            name = prop.get('name')
            value = prop.get('value', '').strip()

            if name and name.startswith('pop:'):
                clean_name = name.split(':', 1)[1]
                
                if clean_name in multi_value_fields and value:
                    pop_properties[clean_name] = [item.strip() for item in value.split('//')]
                    print(f"  - Encontrado (lista): {clean_name} = {pop_properties[clean_name]}")
                else:
                    pop_properties[clean_name] = value
                    if value:
                        print(f"  - Encontrado (texto): {clean_name} = {value}")

        print("\nINFO: Extraindo documentação das tarefas para a Seção III...")
        task_documentations = []
        process_id = participant.get('processRef')
        process_element = root.find(f".//bpmn:process[@id='{process_id}']", namespaces=ns)
        
        if process_element is not None:
            # Busca todos os elementos que possuem uma tag de documentação
            elements_with_docs = process_element.xpath(".//*[bpmn:documentation]", namespaces=ns)
            for elem in elements_with_docs:
                doc_element = elem.find('bpmn:documentation', namespaces=ns)
                # Lida corretamente com o conteúdo HTML dentro da tag de documentação
                doc_text = etree.tostring(doc_element, method='text', encoding='unicode').strip()
                elem_name = elem.get('name')
                if doc_text:
                    task_documentations.append({
                        "elemento": elem_name if elem_name else elem.tag.split('}', 1)[1],
                        "descricao": doc_text
                    })
                    print(f"  - Documentação encontrada para: '{elem_name}'")

        final_data = {
            "propriedades_pop": pop_properties,
            "descricao_processo_atividades": task_documentations
        }

        return final_data

    except Exception as e:
        print(f"ERRO: Ocorreu um erro inesperado durante a análise. Erro: {e}")
        return None

# --- Bloco de Execução ---
if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python parser_bpmn.py /caminho/para/seu/arquivo.bpmn")
        sys.exit(1)

    bpmn_file_path = sys.argv[1]
    
    extracted_data = parse_bpmn_pop(bpmn_file_path)

    if extracted_data:
        print("\n--- DADOS EXTRAÍDOS COM SUCESSO ---")
        print(json.dumps(extracted_data, indent=2, ensure_ascii=False))
        print("---------------------------------")
