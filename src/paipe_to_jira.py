# -*- coding: UTF-8 -*-
import sys
import getpass
import base64
import requests
import xlrd


class Issue:
    def __init__(self, alias, issue_type):
        self.alias = alias
        self.issue_type = issue_type
        self.key = None
        self.fields = {}
        self.links = []
        self.relations_design = []
        self.relations_other = []


class Jira:
    def __init__(self):
        self.login = ''
        self.password = ''
        self.project = ''
        self.board = ''
        url = self.get_url()
        self.get_login()
        self.get_password()
        self.get_project()
        self.get_board()
        self.url = 'http://{0}/rest/api/2'.format(url)
        self.agile = 'http://{0}/rest/greenhopper/1.0'.format(url)

        self.default_component = None
        self.issues = []
        self.jira_issues = []
        self.supported_fields = []
        self.header = {'Content-Type': 'application/json',
                       'Authorization': 'Basic '
                                        + base64.b64encode(self.login
                                                           + ':'
                                                           + self.password)
                       }
        response = self.request('GET', 'issue/createmeta?projectKeys={0}'.format(self.project))
        print
        if response.status_code != 200:
            print('ERROR: {0}'.format(
                'Unable to login to Jira\nMake sure to your login and password are correct\nAnd try logging in using webrowser'))
            sys.exit(1)
        if len(response.json()['projects']) != 1:
            print('ERROR: {0}'.format('Wrong project id'))
            sys.exit(1)

    def request(self, http_method, method, json=None):
        response = requests.request(http_method, u'{0}/{1}'.format(self.url, method),
                                    json=json, headers=self.header, timeout=60)
        return response

    def greenhopper(self, http_method, method, json=None):
        response = requests.request(http_method, u'{0}/{1}'.format(self.agile, method),
                                    json=json, headers=self.header, timeout=60)
        return response

    def get_field(self, name, issue_type):
        response = self.request('GET',
                                'issue/createmeta?projectKeys={0}&issuetypeNames={1}&expand=projects.issuetypes.fields'.format(
                                    self.project, issue_type))
        try:
            fields = response.json()['projects'][0]['issuetypes'][0]['fields']
        except:
            return None
        for k, v in fields.iteritems():
            if v['name'] == name:
                return k

    def add_issue(self, issue):
        fields = issue.fields.copy()
        fields['project'] = {'key': self.project}
        fields['issuetype'] = {'name': issue.issue_type}
        json = {
            'fields': fields
        }
        response = self.request('POST', 'issue/', json, )
        if response.status_code != 201:
            return None
        else:
            return response.json()['key']

    def edit_issue(self, issue):
        fields = issue.fields.copy()
        fields['project'] = {'key': self.project}
        fields['issuetype'] = {'name': issue.issue_type}
        del fields[self.get_field('Sprint', issue.issue_type)]
        json = {
            'fields': fields
        }
        response = self.request('PUT', u'issue/{0}'.format(issue.key), json)
        if response.status_code != 204:
            return None
        else:
            return issue.key

    def update_links(self, issue, i_links, type):
        links = []
        for i in i_links:
            links += [e.key for e in self.issues if i == e.alias and issue.key != e.key]

        for link in links:
            json = {
                'type': {
                    'name': type
                },
                'inwardIssue': {
                    'key': issue.key
                },
                'outwardIssue': {
                    'key': link
                },

            }

            response = self.request('POST', u'issueLink', json)
            if response.status_code != 201:
                print('Error when creating link from {0} to {1}'.format(issue.key, link))

    def get_issues_from_jira(self, issuetype):
        query = u'project = {0} AND \"issuetype\"=\"{1}\"'.format(self.project, issuetype)
        start_at = 0
        max_results = 50
        while True:
            json = {u'jql': query, 'startAt': start_at, 'maxResults': max_results}
            result = self.request('POST', u'search', json=json).json()
            self.jira_issues.extend(result['issues'])

            if start_at + max_results >= result['total']:
                break
            start_at += max_results

    def get_issues_from_paipe_xml(self, filename, issue_type):
        book = xlrd.open_workbook(filename)
        sh = book.sheet_by_index(0)
        if sh.name != u'Paipe - Work Package Export':
            print(u'Warning: Looks like this xls was not imported from Paipe')
            sys.exit(1)
        alias_index = [e.value for e in sh.row(0)].index(u'Alias')
        work_index = [e.value for e in sh.row(0)].index(u'Work Package Responsible')
        title_slogan_index = [e.value for e in sh.row(0)].index(u'Title/Slogan')
        how_to_demo_index = [e.value for e in sh.row(0)].index(u'How to demo')
        relations_index = [e.value for e in sh.row(0)].index(
            u'Relations, Integration Dependencies  (towards other WPs)')
        issue_type_index = [e.value for e in sh.row(0)].index(u'JIRA issue type')
        status_index = [e.value for e in sh.row(0)].index(u'Status')
        sprint_index = [e.value for e in sh.row(0)].index(u'Shipment')
        relations_design_index = [e.value for e in sh.row(0)].index(
            u'Relations, Design Dependencies  (towards other WPs)')
        relations_other_index = [e.value for e in sh.row(0)].index(u'Relations, Other relations (towards other WPs)')

        for i in range(sh.nrows - 1):
            alias = sh.row(i + 1)[alias_index].value
            work = sh.row(i + 1)[work_index].value
            title_slogan = sh.row(i + 1)[title_slogan_index].value
            how_to_demo = sh.row(i + 1)[how_to_demo_index].value
            relations = sh.row(i + 1)[relations_index].value
            issue_type = sh.row(i + 1)[issue_type_index].value
            status = sh.row(i + 1)[status_index].value
            sprint = sh.row(i + 1)[sprint_index].value
            relations_design = sh.row(i + 1)[relations_design_index].value
            relations_other = sh.row(i + 1)[relations_other_index].value

            colors = {u'on schedule': u'ghx-label-6',
                      u'risk': u'ghx-label-2',
                      u'off track': u'ghx-label-9',
                      u'no fill': u'#FFFFFF',
                      u'completed': u'ghx-label-4',
                      u'': u'#FFFFFF'}

            color = colors[status.lower()]


            issue = Issue(alias,issue_type)

            issue.fields['summary'] = u'Anatom {0}: {1}'.format(alias, title_slogan).replace('\r','')\
                                                                                    .replace('\n','')
            issue.fields['components'] = [{'name': c} for c in work.split(';')]
            if sprint.strip() != '':
                months = [['January', 'Jan'],
                          ['February', 'Feb'],
                          ['March', 'Mar'],
                          ['April', 'Apr'],
                          ['May', 'May'],
                          ['June', 'Jun'],
                          ['July', 'Jul'],
                          ['August', 'Aug'],
                          ['September', 'Sep'],
                          ['October', 'Oct'],
                          ['November', 'Nov'],
                          ['December', 'Dec'],
                          ]
                for mm,m in months:
                    sprint = sprint.replace(mm,m)
                if len('{0} {1}'.format(self.project,sprint))>30:
                    print('Error: Sprint name for {0} is longer then 30 characters'.format(issue.alias))
                else:
                    issue.fields[self.get_field('Sprint', issue_type)]= str(self.get_sprint('{0} {1}'.format(self.project,sprint)))

            issue.links = relations.split(';')
            issue.relations_design = relations_design.split(';')
            issue.relations_other = relations_other.split(';')
            if issue_type == 'Epic':
                issue.fields[self.get_field('Epic Name', 'Epic')] = u'{0}: {1}'.format(alias, title_slogan)\
                                                                               .replace('\r', '')\
                                                                               .replace('\n', '')
                if self.get_field('Epic Color', 'Epic') != None:
                    issue.fields[self.get_field('Epic Color', 'Epic')] = color
                issue.fields['description']= u'How to demo: {0}'.format(how_to_demo) + \
                                             chr(10) + \
                                             u'WP Status: {0}'.format(status) + \
                                             chr(10) + \
                                             u'Integration dependencies: {0}'.format(relations) + \
                                             chr(10) + \
                                             u'Design Dependencies: {0}'.format(relations_design) + \
                                             chr(10) + \
                                             u'Other relations: {0}'.format(relations_other)

            elif issue_type == 'Task':
                issue.fields['description']= u'WP Status: {0}'.format(status) + \
                                             chr(10) + \
                                             u'Design Dependencies: {0}'.format(relations_design) + \
                                             chr(10) + \
                                             u'Other relations: {0}'.format(relations_other)
            else:
                print('Error: Unknow issue type for {0}'.format(alias))
            issue.links = relations.split(';')
            issue.relations_design = relations_design.split(';')
            issue.relations_other = relations_other.split(';')

            self.issues.append(issue)

    def get_epic_key(self, epic_name):
        query = u'project = {0} AND \"summary\"~\"Anatom {1}:\"'.format(self.project, epic_name)
        json = {u'jql': query, 'startAt': 0}
        result = self.request('POST', u'search', json=json).json()
        if result['total'] < 1:
            return None
        for issue in result['issues']:
            if issue['fields']['summary'].startswith('Anatom WP {0}:'.format(epic_name)):
                return issue['key']
        return None

    def get_components(self):
        response = jira.request('GET',
                                'issue/createmeta?projectKeys={0}&issuetypeNames=Epic&expand=projects.issuetypes.fields'.format(
                                    self.project))
        return [a['name'] for a in
                response.json()['projects'][0]['issuetypes'][0]['fields']['components']['allowedValues']]

    def get_login(self):
        self.login = raw_input('JIRA login: ')

    def get_password(self):
        self.password = getpass.getpass('JIRA password: ')

    def get_project(self):
        self.project = raw_input('JIRA project: ')

    def get_board(self):
        self.board = raw_input('Project board: ')

    def get_url(self):
        return raw_input('JIRA url: ')

    def add_component(self, name):
        json = {
            "name": name['name'],
            "description": u"This JIRA component was added automatically",
            "project": self.project
        }
        response = jira.request('POST', 'component', json=json)
        return response.status_code

    def get_sprints(self):
        board_id = self.get_board_id()

        response = self.greenhopper('GET', 'sprintquery/{0}?includeFutureSprints=true&includeHistoricSprints=false'.format(board_id))
        if response.status_code != 200:
            print('Error: Can\'t get board with id {0}'.format(board_id))
        sprints = response.json()['sprints']
        return sprints

    def get_board_id(self):
        response = self.greenhopper('GET', 'rapidviews/list')
        if response.status_code != 200:
            raise Exception("Wrong response code for rapidviews/list")
        try:
            board_id = [view['id'] for view in response.json()['views'] if view['name'].strip() == self.board.strip()][0]
        except:
            print('Error: No agile board named {0}'.format(self.board))
            sys.exit(1)
        return board_id

    def create_sprint(self, name):
        response = self.greenhopper('POST', 'sprint/{0}'.format(self.get_board_id()))
        id = response.json()['id']
        response = self.greenhopper('PUT', 'sprint/{0}'.format(id), {'name': name})
        return id

    def get_sprint(self, name):

        sprints = self.get_sprints()
        for sprint in sprints:
            if sprint['name'] == name:
                return sprint['id']
        return self.create_sprint(name)


    def get_project_name(self):
        response = self.request('GET', 'project/{0}'.format(self.project))
        if response.status_code != 200:
            print('Error: Can\'t find project with key={0}'.format(self.project))
        return response.json()['name']


if __name__ == '__main__':
    if len(sys.argv) != 2:
        print('No paipe xls file specified!\nUsage: paipe_to_jira <xls file>')
        sys.exit(1)
    else:
        filename = sys.argv[1]
    try:
        jira = Jira()
    except Exception as e:
        print 'ERROR: {0}'.format(e.message)
        sys.exit(1)

    jira.get_issues_from_paipe_xml(filename, 'Epic')
    jira.issues.sort(key=lambda issue: int(issue.alias.split(' ')[-1]))
    for issue in jira.issues:
        issue.key = jira.get_epic_key(issue.fields['summary'].split(':')[0].split(' ')[-1])
        for c in issue.fields['components']:
            if c['name'] not in jira.get_components():
                if jira.add_component(c) not in [201, 204]:
                    print('Failed to add new component: {0}'.format(c['name']))
                else:
                    print('Added new component: {0}'.format(c))
        if issue.key:
            print(u'Updating {0}'.format(issue.alias))
            if not jira.edit_issue(issue):
                print(u'Failed to update {0}'.format(issue.alias))
        else:
            print(u'Creating {0}'.format(issue.alias))
            issue.key = jira.add_issue(issue)
            if not issue.key:
                print(u'Failed to create {0}'.format(issue.alias))

    for issue in jira.issues:
        print(u'Updating links for {0}'.format(issue.alias))
        jira.update_links(issue, issue.links, 'Parent')
        jira.update_links(issue, issue.relations_design, 'Relates')
        jira.update_links(issue, issue.relations_other, 'Relates')
