# encoding: iso-8859-1

import datetime
import config
from jira import JIRA


class JiraProcess(object):
    """Auditoria de processo Protheus. """

    jira = None
    database = None
    cursor = None

    def __init__(self):
        """Inicializador da classe. """
        self.jiraconnect()
        pass

    def jiraconnect(self):
        """Conexao com o JIRA. """
        '''try:
            JIRA_SERVER = 'http://jiraproducao.XXXXXX.com.br'
            self.jira = JIRA(options={'server': JIRA_SERVER}, oauth={
                'access_token': config.access_token,
                'access_token_secret': config.access_token_secret,
                'consumer_key': config.CONSUMER_KEY,
                'key_cert': config.RSA_KEY})
        except:'''

        options = {'server': 'http://' + config.base_jira + '.com.br',
                   "headers": {
                       'User-Agent': 'CODIGO DE USER AGENT',}
                   }
        try:
            self.jira = JIRA(options, basic_auth=(config.user, config.password))
            print("Conectado em " + config.base_jira)
        except:
            print("Algo errado com a conexão em " + config.base_jira + " | Login: " + config.user)
            exit()
        return self.jira

    def getIssue(self, issue):
        """Retorna a issue do Jira. """
        try:
            issue = self.jira.issue(issue, expand='changelog')
        except:
            print('Issue ' + issue +  ' not exist!')
            issue = ""
        return issue

    def getIssueType(self, issue):
        """ Retorna o tipo de uma issue. """
        issueType = ''
        issue = self.getIssue(issue)
        if issue:
            issueType = issue.fields.issuetype
            issueType = issueType.name
        return issueType

    def getTransition(self, issue):
        """ Retorna a transição de uma issue. """
        transition = []
        issue = self.getIssue(issue)
        if issue:
            changelog = issue.changelog
            for history in changelog.histories:
                for item in history.items:
                    if item.field == 'status':
                        transition.append([item.toString,history.created,history.author.key])

        return transition

    def getListIssues(self, filter):
        """ Retorna a lista de issues contidas no filtro informado. """
        listissues = self.jira.search_issues(filter,maxResults=100)#maxResults=800)
        return listissues


    def getCycle_Queue(self, issue):
        """ Retorna a os campos necessarios para Cycle e Queue """
        issue = self.getIssue(issue)

        #converte os campos para formato data
        created = datetime.datetime.strptime(issue.fields.created, '%Y-%m-%dT%H:%M:%S.%f%z') #Created
        dataini = datetime.datetime.strptime(issue.fields.customfield_17100, '%Y-%m-%dT%H:%M:%S.%f%z') # Data Inicio Planejado
        resolved = datetime.datetime.strptime(issue.fields.resolutiondate, '%Y-%m-%dT%H:%M:%S.%f%z')  # Resolvido

        issues = [created,dataini,resolved,resolved - dataini, dataini - created]  # + cycle , queue
        return issues

    def getLead_Suporte(self, issue):
        """ Retorna a os campos necessarios para Lead Time e Tempo no suporte"""
        issue = self.getIssue(issue)

        # converte os campos para formato data
        created = datetime.datetime.strptime(issue.fields.created, '%Y-%m-%dT%H:%M:%S.%f%z')  # Created
        dataaber = datetime.datetime.strptime(issue.fields.customfield_11043,'%Y-%m-%dT%H:%M:%S.%f%z')  #  Data abertura do ticket
        resolved = datetime.datetime.strptime(issue.fields.resolutiondate, '%Y-%m-%dT%H:%M:%S.%f%z')  # Resolvido

        issues = [created, dataaber, resolved, resolved - dataaber, created - dataaber]  # + Lead , Suporte
        return issues

    def getIssueCreationDate(self, issue):
        """ Retorna a Data e Hora de Criação da Issue. """
        issue = self.getIssue(issue)
        return issue.fields.created


    def getTicketOpenDate(self, issue):
        """ Retorna a Data e Hora de Abertura do Ticket. """
        issue = self.getIssue(issue)
        return issue.fields.customfield_11043

    def getInicioPlanejado(self, issue):
        """ Retorna a Data e Hora de Inicio Planejado. """
        issue = self.getIssue(issue)
        return issue.fields.customfield_17100

    def close(self):
        '''Fecha a conction com o JIRA '''
        session = getattr(self, "jira", None)
        if session is not None:
            try:
                session.close()
                print("JIRA sesssion closed")
            except TypeError:
                # TypeError: "'NoneType' object is not callable"
                # Could still happen here because other references are also
                # in the process to be torn down, see warning section in
                # https://docs.python.org/2/reference/datamodel.html#object.__del__
                print("JIRA session was not closed")
                pass
            self.jira = None
