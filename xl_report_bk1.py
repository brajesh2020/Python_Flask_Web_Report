from flask import Flask,render_template
#from flask import redirect,url_for,request,flash
import xlrd
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime,timedelta,date
from sqlalchemy import func
app=Flask(__name__)
app.debug=True
app.secret_key = 'f54ertsdfg'
app.config['SQLALCHEMY_DATABASE_URI'] = 'mysql://root:@localhost/laravel'

db = SQLAlchemy(app)

class xlreport(db.Model):
	quarter = db.Column(db.String(200));	   year = db.Column(db.String(200))
	month = db.Column(db.String(200)) ;	   closed_date = db.Column(db.String(200))
	ticket_category = db.Column(db.String(200))  ;	   master_ticket = db.Column(db.String(200),primary_key = True)
	duration = db.Column(db.Integer()) ;	   age = db.Column(db.Integer())
	location = db.Column(db.String(200));	   issue_day = db.Column(db.String(200))
	issue_start_time = db.Column(db.String(200)) ;	   issue_end_time = db.Column(db.String(200))
	description = db.Column(db.String(200)) ;	   rca_summary = db.Column(db.String(200))
	tower = db.Column(db.String(200)) ;	   tower1 = db.Column(db.String(200))
	cause_category = db.Column(db.String(200)) ;	   trigger_event = db.Column(db.String(200))
	root_cause = db.Column(db.String(200)) ;	   pr = db.Column(db.String(200))
	mco_category = db.Column(db.String(200));	   location2 = db.Column(db.String(200))
	location3 = db.Column(db.String(200)) ;	   no_of_customers = db.Column(db.String(200))
	no_of_vms = db.Column(db.String(200)) ;	   dc = db.Column(db.String(200))
	tower_v2 = db.Column(db.String(200)) ;	   incident_geo = db.Column(db.String(200))
	column1 = db.Column(db.String(200));	   status = db.Column(db.String(200))
	mco_gap_age = db.Column(db.String(200));	Column2 = db.Column(db.String(200))
	t1_or_t2 = db.Column(db.String(200))

	def __init__(self,quarter,year ,month ,closed_date ,ticket_category ,master_ticket ,duration ,age ,location ,issue_day ,issue_start_time ,
              issue_end_time ,description ,rca_summary ,tower ,tower1 ,cause_category ,trigger_event ,root_cause ,pr ,mco_category ,location2 ,
              location3 ,no_of_customers ,no_of_vms ,dc ,tower_v2 ,incident_geo ,column1 ,status ,mco_gap_age ,Column2 ,t1_or_t2):
		self.quarter =  quarter
		self.year = year
		self.month = month
		self.closed_date = closed_date
		self.ticket_category = ticket_category
		self.master_ticket = master_ticket
		self.duration = duration
		self.age = age
		self.location = location
		self.issue_day = issue_day
		self.issue_start_time = issue_start_time
		self.issue_end_time = issue_end_time
		self.description = description
		self.rca_summary = rca_summary
		self.tower = tower
		self.tower1 = tower1
		self.cause_category = cause_category
		self.trigger_event = trigger_event
		self.root_cause = root_cause
		self.pr = pr
		self.mco_category = mco_category
		self.location2 = location2
		self.location3 = location3
		self.no_of_customers = no_of_customers
		self.no_of_vms = no_of_vms
		self.dc = dc
		self.tower_v2 = tower_v2
		self.incident_geo = incident_geo
		self.column1 = column1
		self.status = status
		self.mco_gap_age = mco_gap_age
		self.Column2 = Column2
		self.t1_or_t2 = t1_or_t2

def from_excel_ordinal(ordinal, _epoch=date(1900, 1, 1)):
    if ordinal > 59:
        ordinal -= 1  # Excel leap year bug, 1900 is not a leap year!
    return _epoch + timedelta(days=ordinal - 1)



def str2date(mydate):
	a= datetime.strptime(mydate,'%Y-%m-%d')
	return a

@app.route('/')
def index():
	workbook = xlrd.open_workbook('sample_data.xlsx')
	sheet = workbook.sheet_by_index(0)
	col_name = sheet.row_values(0)
	#excel_list = []
	for rowx in range(1,sheet.nrows):
		col_value = sheet.row_values(rowx)
		xl_dict = {}
		for i in range(len(col_name)):
			colmn = col_name[i].replace(' ','_')

			if colmn == 'Closed_Date':
				col_value[i] = from_excel_ordinal(col_value[i])
			if colmn == 'Age':
				diff = datetime.now().date()-xl_dict['Closed_Date']
				col_value[i]=diff.days
			xl_dict[colmn]=col_value[i]

		xl_master = xlreport.query.filter_by(master_ticket=xl_dict['Master_Ticket']).first()
		if str(xl_master.master_ticket ) == str(xl_dict['Master_Ticket']):
			db.session.query(xlreport).filter_by(master_ticket=xl_dict['Master_Ticket']).update({"age": xl_dict['Age']})
			db.session.commit()
		else:
			xlreport_table = xlreport(xl_dict['Quarter'],xl_dict['Year'],xl_dict['Month'],xl_dict['Closed_Date'],xl_dict['Ticket_\nCategory'],
                             xl_dict['Master_Ticket'],xl_dict['Duration'],xl_dict['Age'],xl_dict['Location'],xl_dict['Issue_Day_'],
                             xl_dict['Issue_Start_\nTime'],xl_dict['Issue_End_\nTime'],xl_dict['Description'],xl_dict['RCA_Summary'],
                             xl_dict['Tower'],xl_dict['Tower1'],xl_dict['Cause_Category'],xl_dict['Trigger_Event'],xl_dict['Root_Cause'],
                             xl_dict['PR#'],xl_dict['MCO_Category'],xl_dict['Location2'],xl_dict['Location3'],xl_dict['No_of_Customers'],
                             xl_dict['No_of_VMs'],xl_dict['DC'],xl_dict['Tower_V2'],xl_dict['Incident_Geo'],xl_dict['Column1'],xl_dict['status'],
                             xl_dict['MCO_Gap_Age'],xl_dict['Column2'],xl_dict['T1_or_T2_'])
			db.session.add(xlreport_table)
			db.session.commit()
#No of Days since last Multi Customer Outages for CMS
#	age_results = db.session.query(xlreport).order_by(db.asc(xlreport.age)).limit(1)
	age_results = db.session.query(xlreport).filter_by(Column2='CMS').order_by(db.asc(xlreport.age)).limit(1)
	mco_results_network = db.session.query(xlreport).filter_by(Column2='CMS',mco_category='NETWORK').order_by(db.asc(xlreport.age)).limit(1)
	mco_results_storage = db.session.query(xlreport).filter_by(Column2='CMS',mco_category='STORAGE').order_by(db.asc(xlreport.age)).limit(1)

#No of Days since last Multi Customer Outages SL-Infra ( Softlayer / Bluemix )
	mco_results_sl = db.session.query(xlreport).filter(xlreport.mco_category.like('SL%')).order_by(db.asc(xlreport.age)).limit(1)
	mco_results_sl_network = db.session.query(xlreport).filter_by(mco_category='SL_Network').order_by(db.asc(xlreport.age)).limit(1)
#No of Days since Last MCO by Compute Platform
	mco_results_1x = db.session.query(xlreport).filter_by(tower1='Compute',tower_v2='1.X').order_by(db.asc(xlreport.age)).limit(1)
	mco_results_3x = db.session.query(xlreport).filter_by(tower1='Compute',tower_v2='3.X').order_by(db.asc(xlreport.age)).limit(1)
	mco_results_2012 =db.session.query(xlreport).filter_by(tower1='Compute',tower_v2='2012.1').order_by(db.asc(xlreport.age)).limit(1)
	mco_results_2012 =db.session.query(xlreport).filter_by(Cause_Category='Human Error').order_by(db.asc(xlreport.age)).limit(1)
    
# read data into pandas directly from the query `q`
#Data Center Wise Last MCO Incident details
#select Tower_V2 ,master_ticket, min(age) from laravel.xlreport where Tower1 = "Compute" group by Tower_V2;

#Last Incident details for MCOs related to Each Category
#select mco_category,min(age) from laravel.xlreport group by mco_category

#Last MCO details for Host related issue by Platform(1.x,2.x,3.x or 2012.1) for Compute
#select Tower_V2 ,min(age) from laravel.xlreport where Tower1 = "Compute" group by Tower_V2

# my_results = xlreport.query.all()
	return render_template('home.html',age_results=age_results,mco_results_network=mco_results_network,mco_results_storage=mco_results_storage,
                        mco_results_sl=mco_results_sl,mco_results_sl_network=mco_results_sl_network,mco_results_1x=mco_results_1x,
                        mco_results_3x=mco_results_3x,mco_results_2012=mco_results_2012)
if __name__=='__main__':
	db.create_all()
	app.run(debug = True)
