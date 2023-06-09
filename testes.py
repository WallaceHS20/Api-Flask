import pandas as pd
import json
from pandas import json_normalize

teste = [
  {
    "nome": "Maria Souza",
    "idade": 35,
    "email": "maria.souza@example.com",
    "endereco": {
      "rua": "Avenida das Palmeiras",
      "numero": 456,
      "complemento": "Apto 123",
      "bairro": "Centro",
      "cidade": "Rio de Janeiro",
      "estado": "RJ",
      "cep": "20000-000",
      "pais": "Brasil"
    },
    "telefones": [
      {
        "tipo": "residencial",
        "numero": "(21) 1111-1111"
      },
      {
        "tipo": "celular",
        "numero": "(21) 99999-9999"
      }
    ],
    "interesses": [
      "Música",
      "Esportes",
      "Leitura"
    ],
    "ultimaCompra": {
      "produto": "Smartphone",
      "valor": 1500.99,
      "data": "2023-05-20"
    }
  },
  {
    "nome": "João Silva",
    "idade": 28,
    "email": "joao.silva@example.com",
    "endereco": {
      "rua": "Rua dos Pinheiros",
      "numero": 789,
      "complemento": "Casa 45",
      "bairro": "Jardins",
      "cidade": "São Paulo",
      "estado": "SP",
      "cep": "01234-567",
      "pais": "Brasil"
    },
    "telefones": [
      {
        "tipo": "residencial",
        "numero": "(11) 2222-2222"
      },
      {
        "tipo": "celular",
        "numero": "(11) 88888-8888"
      }
    ],
    "interesses": [
      "Esportes",
      "Viagens",
      "Gastronomia"
    ],
    "ultimaCompra": {
      "produto": "Notebook",
      "valor": 3000.50,
      "data": "2023-05-15"
    }
  },
  {
    "nome": "Matheus Santos",
    "idade": 32,
    "email": "matheus.santos@example.com",
    "endereco": {
      "rua": "Avenida dos Bandeirantes",
      "numero": 1234,
      "complemento": "Apto 567",
      "bairro": "Moema",
      "cidade": "São Paulo",
      "estado": "SP",
      "cep": "04567-890",
      "pais": "Brasil"
    },
    "telefones": [
      {
        "tipo": "residencial",
        "numero": "(11) 3333-3333"
      },
      {
        "tipo": "celular",
        "numero": "(11) 77777-7777"
      }
    ],
    "interesses": [
      "Tecnologia",
      "Cinema",
      "Artes"
    ],
    "ultimaCompra": {
      "produto": "Headphones",
      "valor": 120.75,
      "data": "2023-06-01"
    }
  }
]

# Normalizar (flatten) os dados JSON
df_normalized = json_normalize(teste)

# Gerar o arquivo Excel
df_normalized.to_excel("dados.xlsx", index=False)
