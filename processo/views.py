from .models import ProcessoAdm, AndamentoAdm, Auditoria

from django.views.generic.edit import CreateView, UpdateView, DeleteView # Módulo para create, update e delete

from django.views.generic.list import ListView # Módulo para list

from django.views.generic import TemplateView, View

from django.urls import reverse, reverse_lazy # Módulo para reverter para a url definida após ter sucesso na execução

import datetime # Módulo para datas

import os # Módulo para trabalhar com pastas e arquivos

from docx2pdf import convert # Módulo para converter .docx em pdf

from PyPDF2 import PdfMerger # Módulo para mesclar pdf

from django.core.exceptions import ValidationError

import pythoncom

from django.http import HttpResponse


###### VIEW ######
class ProcessoAdmView(TemplateView):
    template_name = 'processos/views/processo_adm_view.html'

    # Função para iterar com os dados do processo
    def get_context_data(self, **kwargs):
        processo_pk = self.kwargs.get('pk') # Pega a PK do processo através da URL  

        context = super().get_context_data(**kwargs)
        context['dados_processo'] = ProcessoAdm.objects.filter(pk=processo_pk) # Filtra os dados do processo através da pk
        return context
    
class AndamentoAdmView(TemplateView):
    template_name ='processos/views/andamento_adm_view.html'

     #  Função para reverter para a url 'andamento-adm-list' passando a pk do processo para conseguir voltar para a tela de lista de andamentos do processo.
    def get_voltar(self, processo_pk):
        return reverse('andamento-adm-list-update', args=[processo_pk])

    # Função para funcionalidade do botão 'voltar'
    # Função para buscar a pk do processo e salvar na variável 'processo_pk', com a funcionalidade do get_context_data envia para o Template o contexto 'cancelar' que recebe a função 'get_cancelar' junto com a variavel 'processo_pk'.
    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)

        andamento_pk = self.kwargs.get('pk') # Pega a PK do andamento ao fazer o update através da URL
        andamento = AndamentoAdm.objects.get(pk=andamento_pk) # Busca o andamento através da PK do andamento
        processo_pk = andamento.processo_id # Busca a PK do processo através do andamento (processo_id é a ForeignKey entre o processo administrativo e o andamento)

        context['voltar'] = self.get_voltar(processo_pk)
        context['dados_andamento'] = AndamentoAdm.objects.filter(pk=andamento_pk) # Filtra os dados do andamento através da pk, para conseguir iterar com os dados do andamento

        return context
    
class MesclarPDFsView(View):
    """
        Realiza a mesclagem dos arquivos pdf com o campo checkbox selecionado, disponibilizando o download do arquivo mesclado direto no navegador, sem alterar os arquivos originais.

        Através de um formulário post enviado por um botão submit, pega os pdfs com o checkbox selecionado através do name do checkbox.

        Através do id passado no value do checkbox, pega o atributo arquivo e realiza um append para o PdfMerger, realiza a mesclagem e disponibiliza o arquivo através do HttpResponse.

    """
    def post(self, request):
        pdf_selecionados = request.POST.getlist('pdf_selecionados')

        # Crie um objeto PdfMerger para mesclar os arquivos PDF
        merger = PdfMerger()

        for pdf_id in pdf_selecionados:
            registro = AndamentoAdm.objects.get(id=pdf_id)
            merger.append(registro.arquivo.path)

        # Crie uma resposta para download do PDF mesclado
        response = HttpResponse(content_type='application/pdf')
        response['Content-Disposition'] = 'attachment; filename="pdf_mesclado.pdf"'

        # Mescle os PDFs e envie a resposta
        merger.write(response)
        merger.close()

        return response

###### CREATE ######
class ProcessoAdmCreate(CreateView):
    model = ProcessoAdm
    template_name = 'processos/creates/processo_adm_create.html'
    fields = ['numero', 'municipio', 'uf', 'data_inicial', 'data_final', 'data_div_ativa', 'valor_atributo', 'valor_multa', 'valor_credito', 'valor_atualizado', 'data_valor_atualizado', 'nome_contribuinte', 'tipo_pessoa', 'documento', 'nome_fantasia', 'email', 'endereco', 'complemento', 'municipio_contribuinte', 'uf_contribuinte', 'cep', 'telefone', 'celular']
    success_url = reverse_lazy('proc-adm-list')

    def form_valid(self, form):
        form.instance.usuario_criador = self.request.user 

        result = super().form_valid(form)
  
        # Registre a operação de criação na auditoria
        Auditoria.objects.create(
            usuario = self.request.user,
            model = ProcessoAdm,
            objeto_id = self.object.id,
            view = ProcessoAdmCreate,
            )
        
        return result

class AndamentoAdmCreate(CreateView):  
    model = AndamentoAdm
    template_name = 'processos/creates/andamento_adm_create.html'
    fields = ['data_andamento', 'tipo_andamento', 'situacao_pagamento', 'valor_pago', 'data_prazo', 'data_recebimento', 'complemento', 'arquivo']
    success_url = reverse_lazy('proc-adm-list')

    def form_valid(self, form):
        """
            A função form_valid() serve para alterar os valores do atributo ou realizar qualquer ação antes que o formulário seja salvo.
        """
        pk_processo = self.kwargs.get('pk')
        form.instance.processo_id = pk_processo

        # Preencher o atributo 'criador_andamento_adm' com o ID do usuário logado.
        form.instance.usuario_criador = self.request.user 

        # Preencher o atributo 'funcionario' com o nome completo do usuário logado.
        form.instance.funcionario = self.request.user.get_full_name()

        pythoncom.CoInitialize() # Para não ocorrer o erro  "Exception Value:(-2147221008, 'CoInitialize não foi chamado.', None, None)" quando utilizado código para converter arquivos

        """
        # Código para conversão do arquivo enviado, de .docx(word) para .pdf
        
        """
        # Antes de salvar o formulário, verifica se um arquivo Word foi enviado
        if 'arquivo' in self.request.FILES:
            arquivo = self.request.FILES['arquivo']
            print(arquivo)
            
            if arquivo.name.endswith('.docx'): # Se o arquivo termina com '.docx'
                # Cria um arquivo temporário para a conversão
                word_temporario = os.path.join('media/Arquivo', arquivo.name)
                with open(word_temporario, 'wb') as arquivo_temporario:
                    for chunk in arquivo.chunks():
                        arquivo_temporario.write(chunk)

                # Converte o arquivo Word para PDF
                pdf_temporario = word_temporario.replace('.docx', '.pdf')
                convert(word_temporario, pdf_temporario)

                # Abra o arquivo PDF convertido e atualize o campo 'arquivo' no formulário
                with open(pdf_temporario, 'rb') as pdf:
                    form.instance.arquivo.save(pdf_temporario, pdf)

                # Certifique-se de que o arquivo Word temporário seja excluído
                os.remove(word_temporario)

                # Certifique-se de que o arquivo PDF temporário seja excluído
                os.remove(pdf_temporario)
                
        pythoncom.CoUninitialize() # Para não ocorrer o erro "Exception Value:(-2147221008, 'CoInitialize não foi chamado.', None, None)" quando utilizado códigos para converter arquivos
      
        result = super().form_valid(form)
        
        # Registre a operação de criação na auditoria
        Auditoria.objects.create(
            usuario = self.request.user,
            model = AndamentoAdm,
            objeto_id = self.object.id,
            view = AndamentoAdmCreate,
            )
        
        return result
    
    # Após realizar o create do andamento com sucesso, reverte para a lista de andamentos do processo 
    def get_success_url(self):
        processo_pk = self.kwargs.get('pk') # Pega a PK do processo através da URL       

        return reverse('andamento-adm-list', args=[processo_pk])
    
    # Função para iterar com os dados do processo na view de create andamento
    def get_context_data(self, **kwargs):
        processo_pk = self.kwargs.get('pk') # Pega a PK do processo através da URL  

        context = super().get_context_data(**kwargs)
        context['dados_processo'] = ProcessoAdm.objects.filter(pk=processo_pk) # Filtra os dados do processo através da pk
        return context
    
###### UPDATE ######
class ProcessoAdmUpdate(UpdateView):
    model = ProcessoAdm
    template_name = 'processos/updates/processo_adm_update.html'
    fields = ['municipio', 'uf', 'data_inicial', 'data_final', 'data_div_ativa', 'valor_atributo', 'valor_multa', 'valor_credito', 'valor_atualizado', 'data_valor_atualizado', 'nome_contribuinte', 'tipo_pessoa', 'documento', 'nome_fantasia', 'email', 'endereco', 'complemento', 'municipio_contribuinte', 'uf_contribuinte', 'cep', 'telefone', 'celular']
    success_url = reverse_lazy('proc-adm-list')

    # Função para iterar com os dados do processo na view de update processo
    def get_context_data(self, **kwargs):
        processo_pk = self.kwargs.get('pk') # Pega a PK do processo através da URL  

        context = super().get_context_data(**kwargs)
        context['dados_processo'] = ProcessoAdm.objects.filter(pk=processo_pk) # Filtra os dados do processo através da pk
        return context
    
    def form_valid(self, form):  
        """
            O código a seguir, verifica se algum dos campos do ProcessoAdm foi alterado durante o update, se ocorreu a alteração, será criado um objeto no modelo Auditoria com os dados de registro e os campos alterados
        """    
        objeto_original = self.get_object()
        objeto_atualizado = form.instance

        campos_alterados = []
        
        if objeto_original.municipio != objeto_atualizado.municipio:
            # campos_alterados.append('municipio')
            campos_alterados.append(f'| Campo: municipio ; Valor Antigo: {objeto_original.municipio} ; Valor Novo: {objeto_atualizado.municipio} |')

        if objeto_original.uf != objeto_atualizado.uf:
            campos_alterados.append(f'| Campo: uf ; Valor Antigo: {objeto_original.uf} ; Valor Novo: {objeto_atualizado.uf} |')

        if objeto_original.data_inicial != objeto_atualizado.data_inicial:
            campos_alterados.append(f'| Campo: data_inicial ; Valor Antigo: {objeto_original.data_inicial} ; Valor Novo: {objeto_atualizado.data_inicial} |')

        if objeto_original.data_final != objeto_atualizado.data_final:
            campos_alterados.append(f'| Campo: data_final ; Valor Antigo: {objeto_original.data_final} ; Valor Novo: {objeto_atualizado.data_final} |')

        if objeto_original.data_div_ativa != objeto_atualizado.data_div_ativa:
            campos_alterados.append(f'| Campo: data_div_ativa ; Valor Antigo: {objeto_original.data_div_ativa} ; Valor Novo: {objeto_atualizado.data_div_ativa} |')

        if objeto_original.valor_atributo != objeto_atualizado.valor_atributo:
            campos_alterados.append(f'| Campo: valor_atributo ; Valor Antigo: {objeto_original.valor_atributo} ; Valor Novo: {objeto_atualizado.valor_atributo} |')

        if objeto_original.valor_multa != objeto_atualizado.valor_multa:
            campos_alterados.append(f'| Campo: valor_multa ; Valor Antigo: {objeto_original.valor_multa} ; Valor Novo: {objeto_atualizado.valor_multa} |')

        if objeto_original.valor_credito != objeto_atualizado.valor_credito:
            campos_alterados.append(f'| Campo: valor_credito ; Valor Antigo: {objeto_original.valor_credito} ; Valor Novo: {objeto_atualizado.valor_credito} |')

        if objeto_original.valor_atualizado != objeto_atualizado.valor_atualizado:
            campos_alterados.append(f'| Campo: valor_atualizado ; Valor Antigo: {objeto_original.valor_atualizado} ; Valor Novo: {objeto_atualizado.valor_atualizado} |')

        if objeto_original.data_valor_atualizado != objeto_atualizado.data_valor_atualizado:
            campos_alterados.append(f'| Campo: data_valor_atualizado ; Valor Antigo: {objeto_original.data_valor_atualizado} ; Valor Novo: {objeto_atualizado.data_valor_atualizado} |')

        if objeto_original.nome_contribuinte != objeto_atualizado.nome_contribuinte:
            campos_alterados.append(f'| Campo: nome_contribuinte ; Valor Antigo: {objeto_original.nome_contribuinte} ; Valor Novo: {objeto_atualizado.nome_contribuinte} |')

        if objeto_original.tipo_pessoa != objeto_atualizado.tipo_pessoa:
            campos_alterados.append(f'| Campo: tipo_pessoa ; Valor Antigo: {objeto_original.tipo_pessoa} ; Valor Novo: {objeto_atualizado.tipo_pessoa} |')

        if objeto_original.documento != objeto_atualizado.documento:
            campos_alterados.append(f'| Campo: documento ; Valor Antigo: {objeto_original.documento} ; Valor Novo: {objeto_atualizado.documento} |')

        if objeto_original.nome_fantasia != objeto_atualizado.nome_fantasia:
            campos_alterados.append(f'| Campo: nome_fantasia ; Valor Antigo: {objeto_original.nome_fantasia} ; Valor Novo: {objeto_atualizado.nome_fantasia} |')

        if objeto_original.email != objeto_atualizado.email:
            campos_alterados.append(f'| Campo: email ; Valor Antigo: {objeto_original.email} ; Valor Novo: {objeto_atualizado.email} |')

        if objeto_original.endereco != objeto_atualizado.endereco:
            campos_alterados.append(f'| Campo: endereco ; Valor Antigo: {objeto_original.endereco} ; Valor Novo: {objeto_atualizado.endereco} |')

        if objeto_original.complemento != objeto_atualizado.complemento:
            campos_alterados.append(f'| Campo: complemento ; Valor Antigo: {objeto_original.complemento} ; Valor Novo: {objeto_atualizado.complemento} |')

        if objeto_original.municipio_contribuinte != objeto_atualizado.municipio_contribuinte:
            campos_alterados.append(f'| Campo: municipio_contribuinte ; Valor Antigo: {objeto_original.municipio_contribuinte} ; Valor Novo: {objeto_atualizado.municipio_contribuinte} |')
            
        if objeto_original.uf_contribuinte != objeto_atualizado.uf_contribuinte:
            campos_alterados.append(f'| Campo: uf_contribuinte ; Valor Antigo: {objeto_original.uf_contribuinte} ; Valor Novo: {objeto_atualizado.uf_contribuinte} |')

        if objeto_original.cep != objeto_atualizado.cep:
            campos_alterados.append(f'| Campo: cep ; Valor Antigo: {objeto_original.cep} ; Valor Novo: {objeto_atualizado.cep} |')

        if objeto_original.telefone != objeto_atualizado.telefone:
            campos_alterados.append(f'| Campo: telefone ; Valor Antigo: {objeto_original.telefone} ; Valor Novo: {objeto_atualizado.telefone} |')

        if objeto_original.celular != objeto_atualizado.celular:
            campos_alterados.append(f'| Campo: celular ; Valor Antigo: {objeto_original.celular} ; Valor Novo: {objeto_atualizado.celular} |')

        if campos_alterados:
            # Registra a operação de alteração na auditoria
            Auditoria.objects.create(
                usuario = self.request.user,
                model = ProcessoAdm,
                objeto_id = self.object.id,
                view = ProcessoAdmUpdate,
                campos_alterados = campos_alterados,              
            )

        return super().form_valid(form)

class AndamentoAdmUpdate(UpdateView):
    model = AndamentoAdm
    template_name = 'processos/updates/andamento_adm_update.html'
    fields = ['data_andamento', 'tipo_andamento', 'situacao_pagamento','valor_pago', 'data_prazo', 'data_recebimento', 'complemento', 'arquivo']

    # Após realizar o update com sucesso, reverte para a lista de andamentos do processo
    def get_success_url(self):
        andamento_pk = self.kwargs.get('pk') # Pega a PK do andamento ao fazer o update através da URL
        andamento = AndamentoAdm.objects.get(pk=andamento_pk) # Busca o andamento através da PK do andamento
        processo_pk = andamento.processo_id # Busca a PK do processo através do andamento (processo_id é a ForeignKey entre o processo administrativo e o andamento)

        return reverse('andamento-adm-list-update', args=[processo_pk]) # URL da lista de andamentos + pk do processo 
    
    #  Função para reverter para a url 'andamento-adm-list' passando a pk do processo para conseguir voltar para a tela de lista de andamentos do processo.
    def get_cancelar(self, processo_pk):
        return reverse('andamento-adm-list-update', args=[processo_pk])

    # Função para funcionalidade do botão 'Cancelar'
    # Função para buscar a pk do processo e salvar na variável 'processo_pk', com a funcionalidade do get_context_data envia para o Template o contexto 'cancelar' que recebe a função 'get_cancelar' junto com a variavel 'processo_pk'.
    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)

        andamento_pk = self.kwargs.get('pk') # Pega a PK do andamento ao fazer o update através da URL
        andamento = AndamentoAdm.objects.get(pk=andamento_pk) # Busca o andamento através da PK do andamento
        processo_pk = andamento.processo_id # Busca a PK do processo através do andamento (processo_id é a ForeignKey entre o processo administrativo e o andamento)

        context['cancelar'] = self.get_cancelar(processo_pk)
        context['dados_andamento'] = AndamentoAdm.objects.filter(pk=andamento_pk) # Filtra os dados do andamento através da pk, para conseguir iterar com os dados do andamento

        return context
    
    def form_valid(self, form):
        """
            A função form_valid() serve para alterar os valores do atributo ou realizar qualquer ação antes que o formulário seja salvo.
        """
        objeto_original = self.get_object()
        objeto_atualizado = form.instance

        pythoncom.CoInitialize() # Para não ocorrer o erro  "Exception Value:(-2147221008, 'CoInitialize não foi chamado.', None, None)" quando utilizado código para converter arquivos
        # Código para conversão do arquivo enviado, de .docx(word) para .pdf
        # Antes de salvar o formulário, verifica se um arquivo Word foi enviado
        if 'arquivo' in self.request.FILES:
            arquivo = self.request.FILES['arquivo']
            print(arquivo)
            
            if arquivo.name.endswith('.docx'): # Se o arquivo termina com '.docx'
                # Cria um arquivo temporário para a conversão
                word_temporario = os.path.join('media/Arquivo', arquivo.name)
                with open(word_temporario, 'wb') as arquivo_temporario:
                    for chunk in arquivo.chunks():
                        arquivo_temporario.write(chunk)

                # Converte o arquivo Word para PDF
                pdf_temporario = word_temporario.replace('.docx', '.pdf')
                convert(word_temporario, pdf_temporario)

                # Abra o arquivo PDF convertido e atualize o campo 'arquivo' no formulário
                with open(pdf_temporario, 'rb') as pdf:
                    form.instance.arquivo.save(pdf_temporario, pdf)

                # Certifique-se de que o arquivo Word temporário seja excluído
                os.remove(word_temporario)

                # Certifique-se de que o arquivo PDF temporário seja excluído
                os.remove(pdf_temporario)
                
                pythoncom.CoUninitialize() # Para não ocorrer o erro "Exception Value:(-2147221008, 'CoInitialize não foi chamado.', None, None)"

        # result = super().form_valid(form)

        campos_alterados = []

        if objeto_original.data_andamento != objeto_atualizado.data_andamento:
            campos_alterados.append(f'| Campo: data_andamento ; Valor Antigo: {objeto_original.data_andamento} ; Valor Novo: {objeto_atualizado.data_andamento} |')

        if objeto_original.tipo_andamento != objeto_atualizado.tipo_andamento:
            campos_alterados.append(f'| Campo: tipo_andamento ; Valor Antigo: {objeto_original.tipo_andamento} ; Valor Novo: {objeto_atualizado.tipo_andamento} |')

        if objeto_original.situacao_pagamento != objeto_atualizado.situacao_pagamento:
            campos_alterados.append(f'| Campo: situacao_pagamento ; Valor Antigo: {objeto_original.situacao_pagamento} ; Valor Novo: {objeto_atualizado.situacao_pagamento} |')

        if objeto_original.valor_pago != objeto_atualizado.valor_pago:
            campos_alterados.append(f'| Campo: valor_pago ; Valor Antigo: {objeto_original.valor_pago} ; Valor Novo: {objeto_atualizado.valor_pago} |')

        if objeto_original.data_prazo != objeto_atualizado.data_prazo:
            campos_alterados.append(f'| Campo: data_prazo ; Valor Antigo: {objeto_original.data_prazo} ; Valor Novo: {objeto_atualizado.data_prazo} |')

        if objeto_original.data_recebimento != objeto_atualizado.data_recebimento:
            campos_alterados.append(f'| Campo: data_recebimento ; Valor Antigo: {objeto_original.data_recebimento} ; Valor Novo: {objeto_atualizado.data_recebimento} |')
        
        if objeto_original.complemento != objeto_atualizado.complemento:
            campos_alterados.append(f'| Campo: complemento ; Valor Antigo: {objeto_original.complemento} ; Valor Novo: {objeto_atualizado.complemento} |')

        if objeto_original.arquivo != objeto_atualizado.arquivo:
            campos_alterados.append(f'| Campo: arquivo ; Valor Antigo: {objeto_original.arquivo} ; Valor Novo: {objeto_atualizado.arquivo} |')
        
        if campos_alterados:
            # Registra a operação de update na auditoria
            Auditoria.objects.create(
                usuario = self.request.user,
                model = AndamentoAdm,
                objeto_id = self.object.id,
                view = AndamentoAdmUpdate,
                campos_alterados = campos_alterados,
                )
            
        return super().form_valid(form)

###### DELETE ######
class ProcessoAdmDelete(DeleteView):
    model = ProcessoAdm
    template_name = 'processos/deletes/processo_adm_delete.html'
    success_url = reverse_lazy('proc-adm-list')

    def form_valid(self, form):        
        # Registre a operação de criação na auditoria
        Auditoria.objects.create(
            usuario = self.request.user,
            model = ProcessoAdm,
            objeto_id = self.object.id,
            view = ProcessoAdmDelete,
            )
        
        return super().form_valid(form)
    
    # Função para iterar com os dados do processo na view de delete processo
    def get_context_data(self, **kwargs):
        processo_pk = self.kwargs.get('pk') # Pega a PK do processo através da URL  

        context = super().get_context_data(**kwargs)
        context['dados_processo'] = ProcessoAdm.objects.filter(pk=processo_pk) # Filtra os dados do processo através da pk
        return context 


class AndamentoAdmDelete(DeleteView):
    model = AndamentoAdm
    template_name = 'processos/deletes/andamento_adm_delete.html'

    def form_valid(self, form):        
        # Registre a operação de criação na auditoria
        Auditoria.objects.create(
            usuario = self.request.user,
            model = AndamentoAdm,
            objeto_id = self.object.id,
            view = AndamentoAdmDelete,
            )
        
        return super().form_valid(form)

    # Após realizar o delete com sucesso, reverte para a lista de andamentos do processo
    def get_success_url(self):
        andamento_pk = self.kwargs.get('pk') # Pega a PK do andamento ao fazer o update através da URL
        andamento = AndamentoAdm.objects.get(pk=andamento_pk) # Busca o andamento através da PK do andamento
        processo_pk = andamento.processo_id # Busca a PK do processo através do andamento (processo_id é a ForeignKey entre o processo administrativo e o andamento)

        return reverse('andamento-adm-list-update', args=[processo_pk]) # URL da lista de andamentos + pk do processo
    
    # Função para iterar com os dados do andamento na view de delete andamento
    def get_context_data(self, **kwargs):
        andamento_pk = self.kwargs.get('pk') # Pega a PK do andamento através da URL  

        context = super().get_context_data(**kwargs)
        context['dados_andamento'] = AndamentoAdm.objects.filter(pk=andamento_pk) # Filtra os dados do andamento através da pk
        return context

###### LIST ######
class ProcessoAdmList(ListView):
    model = ProcessoAdm
    template_name = 'processos/lists/processo_adm_list.html'

    def get_context_data(self, **kwargs):
        """
            
        """
        processos = ProcessoAdm.objects.all()

        for processo in processos:

            processos_id = ProcessoAdm.objects.get(pk=processo.id)

            andamentos = processos_id.andamentoadm_set.all()

            context = super().get_context_data(**kwargs)
            context['dados_andamento'] = andamentos
        
            return context

class AndamentoAdmList(ListView):
    model = ProcessoAdm
    template_name = 'processos/lists/andamento_adm_list.html'
    
    def get_queryset(self):
        pk_processo = self.kwargs.get('pk') # Pega a pk(primary key) da URL, pk do processo
        
        processo = ProcessoAdm.objects.get(pk=pk_processo)  # Pega o processo que possui a pk recebida (pk é a primary key do processo)
        andamentos = processo.andamentoadm_set.all()  # Pega todos os atributos do andamento
    
        return andamentos
    
    # Função para iterar com os dados do processo na lista de andamentos
    def get_context_data(self, **kwargs):
        processo_pk = self.kwargs.get('pk') # Pega a PK do processo através da URL  

        context = super().get_context_data(**kwargs)
        context['dados_processo'] = ProcessoAdm.objects.filter(pk=processo_pk) # Filtra os dados do processo através da pk
        return context

class AndamentoAdmListUpdate(ListView):
    model = ProcessoAdm
    template_name = 'processos/lists/andamento_adm_list_update.html'
    
    def get_queryset(self):
        pk_processo = self.kwargs.get('pk') # Pega a pk(primary key) da URL, pk do processo
        
        processo = ProcessoAdm.objects.get(pk=pk_processo)  # Pega o processo que possui a pk recebida (pk é a primary key do processo)
        andamentos = processo.andamentoadm_set.all()  # Pega todos os atributos do andamento
    
        return andamentos
    
    # Função para iterar com os dados do processo na lista de andamentos
    def get_context_data(self, **kwargs):
        processo_pk = self.kwargs.get('pk') # Pega a PK do processo através da URL  

        context = super().get_context_data(**kwargs)
        context['dados_processo'] = ProcessoAdm.objects.filter(pk=processo_pk) # Filtra os dados do processo através da pk
        return context

    

