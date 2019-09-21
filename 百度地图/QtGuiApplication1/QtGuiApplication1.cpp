#include "QtGuiApplication1.h"

double PI = 3.1415926535897932384626;
double ee = 0.00669342162296594323;
double a = 6378245.0;
double x_PI = 3.14159265358979324 * 3000.0 / 180.0;

QtGuiApplication1::QtGuiApplication1(QWidget *parent)
	: QDialog(parent)
{
	Qt::WindowFlags ture = Qt::Dialog;
	ture |= Qt::WindowMinimizeButtonHint;
	ture |= Qt::WindowMaximizeButtonHint;
	ture |= Qt::WindowCloseButtonHint;
	setWindowFlags(ture);
	

	m_gridLayoyt = new QGridLayout(this);

	m_excel = new QAxObject(this);
	if (!m_excel->setControl("Excel.Application"))
	{
		QMessageBox::information(NULL, "warning", "Connected Excel.Application failed");
		CoUninitialize();
	}

	m_button = new QPushButton;
	m_button->setText(tr("open xlsx"));

	m_view = new QWebEngineView;

	m_gridLayoyt->addWidget(m_view, 0, 0);
	m_gridLayoyt->addWidget(m_button, 0, 1);

	this->setLayout(m_gridLayoyt);

	connect(m_button, &QPushButton::clicked,
		this, &QtGuiApplication1::onPushButtn);

	connect(m_view, &QWebEngineView::loadFinished,
		this, &QtGuiApplication1::onLoadFinished);

}

void QtGuiApplication1::onPushButtn()
{
	
	
	QString file_path = QFileDialog::getOpenFileName(this, "file path", "./");
	if (file_path.isEmpty())
	{
		QMessageBox::information(NULL, "warning", "File path is empty");
		return;
	}

	qDebug() << file_path << endl;


	openFile(file_path);

	makeMap();

}

void QtGuiApplication1::onLoadFinished(bool a)
{
	QJsonArray num_json, num2_json;                       //声明QJsonArray
	QJsonDocument num_document, num2_document;    //将QJsonArray改为QJsonDocument类
	QByteArray num_byteArray, num2_byteArray;      //

	int i = 0;
	for (auto iter = m_data.begin();iter != m_data.end();iter++)
	{
		i++;
		if (i == 5)
		{
			i = 0;
			num_json.append(iter->x);
			num2_json.append(iter->y);
		}
		else
			continue;
	}

	num_document.setArray(num_json);
	num2_document.setArray(num2_json);
	num_byteArray = num_document.toJson(QJsonDocument::Compact);
	num2_byteArray = num2_document.toJson(QJsonDocument::Compact);
	QString numJson(num_byteArray);             //再转为QString
	QString num2Json(num2_byteArray);            

												  
	
	QString cmd = QString("showarray(\"%1\",\"%2\")").arg(numJson).arg(num2Json);
	m_view->page()->runJavaScript(cmd);          //传给javascript

	
}

void QtGuiApplication1::openFile(QString filename)
{
	QAxObject *workbooks = NULL;
	QAxObject *workbook = NULL;

	m_excel->dynamicCall("SetVisible(bool)", false);
	workbooks = m_excel->querySubObject("WorkBooks");
	workbook = workbooks->querySubObject("Open(QString, QVariant)", filename);

	QAxObject * worksheet = workbook->querySubObject("WorkSheets(int)", 1);
	QAxObject * usedrange = worksheet->querySubObject("UsedRange");
	QAxObject * rows = usedrange->querySubObject("Rows");
	QAxObject * columns = usedrange->querySubObject("Columns");
	int intRows = rows->property("Count").toInt();
	int intCols = columns->property("Count").toInt();

	// 载入数据，这里读取B2:C最后
	QString Range = "B2:C" + QString::number(intRows);
	QAxObject *allEnvData = worksheet->querySubObject("Range(QString)", Range);
	QVariant allEnvDataQVariant = allEnvData->property("Value");
	QVariantList allEnvDataList = allEnvDataQVariant.toList();

	m_data.clear();
	for (int i = 0; i < intRows - 1; i++)
	{
		QVariantList allEnvDataList_i = allEnvDataList[i].toList();
		
		point_ts tmp_data;
		
		tmp_data.x = allEnvDataList_i[0].toDouble();
		tmp_data.y = allEnvDataList_i[1].toDouble();
		point_tsbaidu sss = wgs84tobd09(tmp_data.x, tmp_data.y);
		m_data.push_back(sss);
	}

	workbooks->dynamicCall("Close()");
	
}

void QtGuiApplication1::makeMap()
{
	QString path = QDir::currentPath();
	QUrl url(path.append("/test.html"));

	m_view->load(url);
}

point_tsbaidu QtGuiApplication1::wgs84tobd09(double lng, double lat)
{
	double dlat = transformlat(lng - 105.0, lat - 35.0);
	double dlng = transformlng(lng - 105.0, lat - 35.0);
	double radlat = lat / 180.0 * PI;
	double magic = sin(radlat);
	magic = 1 - ee * magic * magic;
	double sqrtmagic = sqrt(magic);
	dlat = (dlat * 180.0) / ((a * (1 - ee)) / (magic * sqrtmagic) * PI);
	dlng = (dlng * 180.0) / (a / sqrtmagic * cos(radlat) * PI);
	double mglat = lat + dlat;
	double mglng = lng + dlng;

	//第二次转换
	double z = sqrt(mglng * mglng + mglat * mglat) + 0.00002 * sin(mglat * x_PI);
	double theta = atan2(mglat, mglng) + 0.000003 * cos(mglng * x_PI);
	double bd_lng = z * cos(theta) + 0.0065;
	double bd_lat = z * sin(theta) + 0.006;
	
	point_tsbaidu pointmp;
	pointmp.x = bd_lng;
	pointmp.y = bd_lat;

	return pointmp;
}

double QtGuiApplication1::transformlat(double lng, double lat)
{
	double ret = -100.0 + 2.0 * lng + 3.0 * lat + 0.2 * lat * lat + 0.1 * lng * lat + 0.2 * sqrt(abs(lng));
	ret += (20.0 * sin(6.0 * lng * PI) + 20.0 * sin(2.0 * lng * PI)) * 2.0 / 3.0;
	ret += (20.0 * sin(lat * PI) + 40.0 * sin(lat / 3.0 * PI)) * 2.0 / 3.0;
	ret += (160.0 * sin(lat / 12.0 * PI) + 320 * sin(lat * PI / 30.0)) * 2.0 / 3.0;
	return ret;
}

double QtGuiApplication1::transformlng(double lng, double lat)
{
	double ret = 300.0 + lng + 2.0 * lat + 0.1 * lng * lng + 0.1 * lng * lat + 0.1 * sqrt(abs(lng));
	ret += (20.0 * sin(6.0 * lng * PI) + 20.0 * sin(2.0 * lng * PI)) * 2.0 / 3.0;
	ret += (20.0 * sin(lng * PI) + 40.0 * sin(lng / 3.0 * PI)) * 2.0 / 3.0;
	ret += (150.0 * sin(lng / 12.0 * PI) + 300.0 * sin(lng / 30.0 * PI)) * 2.0 / 3.0;
	return ret;
}