#pragma once

#include <QtWidgets/QDialog>
#include <QtWebEngineWidgets/QtWebEngineWidgets>
#include <ActiveQt/QAxObject>
#include <QPushButton>
#include <QGridLayout>
#include <QString>
#include <QMessageBox>
#include <list>
#include <QDebug>
#include <cmath>

struct point_ts
{
	double x;
	double y;
};

struct point_tsbaidu
{
	double x;
	double y;
};



class QtGuiApplication1 : public QDialog
{
	Q_OBJECT

public:
	QtGuiApplication1(QWidget *parent = Q_NULLPTR);

public:
	void openFile(QString filename);
	void makeMap();

	point_tsbaidu wgs84tobd09(double lng, double lat);
	double transformlat(double lng, double lat);
	double transformlng(double lng, double lat);
private slots:
	void onPushButtn();
	void onLoadFinished(bool a);

private:
	QAxObject*          m_excel;
	QString             m_url;
	QPushButton*		m_button;
	QWebEngineView*		m_view;
	QGridLayout*		m_gridLayoyt;
	std::list<point_tsbaidu> m_data;
};
