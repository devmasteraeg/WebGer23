from django.db import models

from django.contrib.auth import get_user_model
from django.contrib.auth.models import User

class Auditoria(models.Model): #Logs de Create, Update, Delete
    usuario = models.ForeignKey(User, on_delete=models.SET_NULL, null=True)
    objeto_id = models.PositiveIntegerField()
    tipo_objeto = models.CharField(max_length=255) # processo administrativo, processo judiciario, andamento
    view = models.CharField(max_length=255)  # "create", "update", "delete"
    acao = models.CharField(max_length=255) # "create", "update", "delete"
    processo = models.CharField(max_length=255, blank=True, null=True)
    andamento = models.CharField(max_length=255, blank=True, null=True)
    campos_alterados = models.CharField(max_length=1000, blank=True, null=True)
    data_hora = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ('-data_hora',)

class Base(models.Model): # Classe base, será herdada pelas outras classes - HERANÇA
    data_criacao = models.DateTimeField('data_criação', auto_now_add=True)
    usuario_criador = models.ForeignKey(get_user_model(), verbose_name='Usuário Criador', on_delete=models.SET_NULL, null=True)
    data_alteracao = models.DateTimeField('Alterado', auto_now=True)
    ativo = models.BooleanField('Ativo', default=True) # Ativo: 1 (True), Inativo: 0 (False)

    class Meta:  # Classe abstrata, não será instânciada.
        abstract = True

class ProcessoAdm(Base):

    tipo_pessoas = (
        ('Física', 'Física'), ('Jurídica', 'Jurídica'),
    )

    numero = models.CharField(unique=True, verbose_name='N°', max_length=10) # Número do processo
    municipio = models.CharField(max_length=50, verbose_name='Município') # Município
    uf = models.CharField(max_length=2) # UF 
    data_inicial = models.DateField(blank=True, null=True) # Data Inicial do Período do processo
    data_final = models.DateField(blank=True, null=True) # Data final do Período do processo
    data_div_ativa = models.DateField(blank=True, null=True) # Data dívida ativa
    valor_atributo = models.DecimalField(decimal_places=2, max_digits=11, blank=True, null=True)  # Valor do atributo
    valor_multa = models.DecimalField(decimal_places=2, max_digits=11, blank=True, null=True) # Valor da multa
    valor_credito = models.DecimalField(decimal_places=2, max_digits=11, blank=True, null=True) # Valor do crédito
    valor_atualizado = models.DecimalField(decimal_places=2, max_digits=11, blank=True, null=True) # Valor do atualizado
    data_valor_atualizado = models.DateField(blank=True, null=True) # Data valor atualizado
    nome_contribuinte = models.CharField(max_length=50)  # Nome / Razão Social
    tipo_pessoa = models.CharField(max_length=50, choices=tipo_pessoas) # Física / Jurídica # Utiliza choices(escolhas) para selecionar o tipo de pessoas
    documento = models.CharField(max_length=20, verbose_name='CPF/CNPJ', unique=True) # CPF / CNPJ
    nome_fantasia = models.CharField(max_length=50, blank=True, null=True) # Nome Fantasia
    email = models.EmailField(max_length=50, blank=True, null=True) # E-mail
    endereco = models.CharField(max_length=150) # Rua
    complemento = models.CharField(max_length=50, blank=True, null=True) # Complemento
    municipio_contribuinte = models.CharField(max_length=50, blank=True, null=True) # Município Contribuinte
    uf_contribuinte = models.CharField(max_length=2, verbose_name='UF', blank=True, null=True) # UF Contribuinte
    cep = models.CharField(max_length=11, blank=True, null=True) # CEP
    telefone = models.CharField(max_length=20, blank=True, null=True) # Telefone
    celular = models.CharField(max_length=20, blank=True, null=True) # Celular

    def __str__(self):
        return f'{self.numero}'

class TipoAndamentoAdm(Base):
    tipo_andamento = models.CharField(max_length=100, verbose_name='Tipo de Andamento')

    def __str__(self):
        return self.tipo_andamento
    
class AndamentoAdm(Base):
    # andamentos = (
    # ('Abertura', 'Abertura'), ('Parecer Fiscal', 'Parecer Fiscal'), ('Decisão 1ª Instância', 'Decisão 1ª Instância'), ('Suspenso Para Fiscalização Futura', 'Suspenso Para Fiscalização Futura'), ('Auto de Infração e Termo de Intimação - AITI.', 'Auto de Infração e Termo de Intimação - AITI.'), ('Termo de Intimação Fiscal - TIF.-tif.', 'Termo de Intimação Fiscal - TIF.'), ('Decisão de 2ª Instância', 'Decisão de 2ª Instância'), ('Cobrança de Documentação', 'Cobrança de Documentação'), ('Recurso Voluntário', 'Recurso Voluntário'), ('Fim do Contrato com a Assessoria', 'Fim do Contrato com a Assessoria'), ('Manifestação', 'Manifestação'), ('Recebimento do AR', 'Recebimento do AR'), ('Despacho', 'Despacho'), ('Aguardando Pagamento', 'Aguardando Pagamento'), ('Apresentação de Documentação para Análise', 'Apresentação de Documentação para Análise'), ('Aguardando AR', 'Aguardando AR'), ('Ofício', 'Ofício'), ('Revelia', 'Revelia'), ('Execução', 'Execução'), ('Confissão de Dívida (Parcelamento)', 'Confissão de Dívida (Parcelamento)'), ('Reenvio de Documento', 'Reenvio de Documento'), ('Parecer Juridico', 'Parecer Juridico'), ('Certidão', 'Certidão'), ('Encaminhado', 'Encaminhado'), ('Encerrado', 'Encerrado'),
    # )

    # Choices
    situacao = (
        ('Sem Pagamento', 'Sem Pagamento'), ('Com Pagamento', 'Com Pagamento'),
    )

    processo = models.ForeignKey(ProcessoAdm, on_delete=models.CASCADE) # Relacionamento 'One to Many' (um para muitos)
    data_andamento = models.DateField(verbose_name='Data do Andamento')
    tipo_andamento = models.ForeignKey(TipoAndamentoAdm, on_delete=models.SET_NULL, null=True) 
    situacao_pagamento = models.CharField(max_length=100, choices=situacao, blank=True, null=True) 
    valor_pago = models.DecimalField(decimal_places=2, max_digits=11, blank=True, null=True)
    data_prazo = models.DateField(blank=True, null=True)
    funcionario = models.CharField(max_length=50, blank=True, null=True)
    data_recebimento = models.DateField(blank=True, null=True)
    complemento = models.CharField(max_length=150, blank=True, null=True)
    arquivo = models.FileField(upload_to='Arquivo/', verbose_name='Arquivo', blank=True) 

    def __str__(self):
        return f'Processo: {self.processo} Andamento: {self.tipo_andamento} Arquivo: {self.arquivo} Ativo: {self.ativo}'

