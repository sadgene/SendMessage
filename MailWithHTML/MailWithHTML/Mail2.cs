using System;
using System.Net;
using System.Net.Mail;
using System.Windows.Forms;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Data.Common;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq;
using System.Data.Linq;
using System.Data.Linq.Mapping;
using System.Xml;
using System.IO;
using CsSQL;
using L2SApp;

public class Mail
{

	static void Main(string[] args)
	{
		DateTime nowdate = DateTime.Now;

		string log = String.Empty;
		string mailadress = String.Empty;
		string dynamiclog = String.Empty;//динамика
		string captionname;
		string nameforadmin; //для админа Менеджеры РКП
		string delay;//отставание		
		string userdelay;//отставание для отправки пользователю
		string overdue;//Просроченные проекты
		string useroverdue;//Просроченные проекты отправка пользователю
		string maxdelay;//отставание более 50
		string usermaxdelay;//отставание пользователю более 50%

		int i = 1;
		int num;
		int numoverdue;


		string datasource = @"server";
		string database = "t_Project";
		string username = "пользователь";
		string password = "пароль";
		string connString = @"Data Source=" + datasource + ";Initial Catalog=" + database + ";Persist Security Info= no ;User ID = " + username + "; Password = " + password;



		DataContext db = new DataContext(connString);


		var Week = db.GetTable<PersentForWeeks>().Where(s => s.MAX_DELTA_PROGRESS < 0).Where(s => s.MAX_DELTA_PROGRESS >= -100).GroupBy(a => a.USER_NAME);





		foreach (var u in Week)
		{
			userdelay = String.Empty;
			usermaxdelay = String.Empty;
			delay = String.Empty;
			maxdelay = String.Empty;
			overdue = String.Empty; ;//Просроченные проекты
			useroverdue = String.Empty;
			captionname = String.Empty;
			nameforadmin = String.Empty;
			captionname += "<p><i>Здравствуйте, " + u.Key.ToString() + " !</i></p>";
			nameforadmin += "<b>Менеджер РКП: " + u.Key.ToString()+"</b>";
			mailadress = String.Empty;
			num = 1;
			numoverdue = 1;
		
			foreach (var c in u)
			{


                
                if (mailadress == String.Empty)
                {
                    if (c.USER_MAIL != String.Empty)
                        mailadress = c.USER_MAIL;
                    else
                        continue;
                }

				if (c.PROJ_REL_CAPTION == null)
				{
					c.PROJ_REL_CAPTION = "Не в рамках договора";

				}

				if (c.I_PLAN_PROGRESS == null)
				{
					c.I_PLAN_PROGRESS = 0;
				}

				if (c.I_FACT_PROGRESS == null)
				{
					c.I_FACT_PROGRESS = 0;
				}

				if (c.MIN_DELTA_PROGRESS == 0)
				{
					if (c.MAX_DELTA_PROGRESS < 0 && c.MAX_DELTA_PROGRESS > -10)
					{
						
						delay += "<p> В период с " +
							c.MIN_PARCE_DATE.ToShortDateString() +
							" по " +
							c.MAX_PARCE_DATE.ToShortDateString() +
							" отставание по проекту<br>" +
							" №" +
							c.PROJ_REL_CAPTION + "<br>" +
							"<a href=\"ips://object/" + c.I_PROJ_ID.ToString() + "\">" + c.I_PROJ_CAPTION + "</a> " + //поменял на стринг							
							" увеличилось на " +
							Convert.ToString(((c.MAX_DELTA_PROGRESS) - (c.MIN_DELTA_PROGRESS)) * -1) +
							"%<br> " +
							"Отставание составляет: <font color=\"green\">" +
							Convert.ToString(c.MAX_DELTA_PROGRESS * -1) +
							"</font>%</p>";

						userdelay += "<p> В период с " +
							c.MIN_PARCE_DATE.ToShortDateString() +
							" по " +
							c.MAX_PARCE_DATE.ToShortDateString() +
							" отставание по проекту<br>" +
							" №" +
							c.PROJ_REL_CAPTION + "<br>" +
							"<a href=\"ips://object/" + c.I_PROJ_ID.ToString() + "\">" + c.I_PROJ_CAPTION + "</a> " + //поменял на стринг							
							" увеличилось на " +
							Convert.ToString(((c.MAX_DELTA_PROGRESS) - (c.MIN_DELTA_PROGRESS)) * -1) +
							"%<br> " +
							"Отставание составляет: <font color=\"green\">" +
							Convert.ToString(c.MAX_DELTA_PROGRESS * -1) +
							"</font>%</p>";
						
					}

					if (c.MAX_DELTA_PROGRESS < -10 && c.MAX_DELTA_PROGRESS > -30)
					{
						delay += "<p> В период с " +
							c.MIN_PARCE_DATE.ToShortDateString() +
							" по " +
							c.MAX_PARCE_DATE.ToShortDateString() +
							" отставание по проекту<br>" +
							" №" +
							c.PROJ_REL_CAPTION + "<br>" +
							"<a href=\"ips://object/" + c.I_PROJ_ID.ToString() + "\">" + c.I_PROJ_CAPTION + "</a> " + //поменял на стринг							
							" увеличилось на " +
							Convert.ToString(((c.MAX_DELTA_PROGRESS) - (c.MIN_DELTA_PROGRESS)) * -1) +
							"%<br> " +
							"Отставание составляет: <font color=\"sandybrown\">" +
							Convert.ToString(c.MAX_DELTA_PROGRESS * -1) +
							"</font>%</p>";

						userdelay += "<p> В период с " +
							c.MIN_PARCE_DATE.ToShortDateString() +
							" по " +
							c.MAX_PARCE_DATE.ToShortDateString() +
							" отставание по проекту<br>" +
							" №" +
							c.PROJ_REL_CAPTION + "<br>" +
							"<a href=\"ips://object/" + c.I_PROJ_ID.ToString() + "\">" + c.I_PROJ_CAPTION + "</a> " + //поменял на стринг							
							" увеличилось на " +
							Convert.ToString(((c.MAX_DELTA_PROGRESS) - (c.MIN_DELTA_PROGRESS)) * -1) +
							"%<br> " +
							"Отставание составляет: <font color=\"sandybrown\">" +
							Convert.ToString(c.MAX_DELTA_PROGRESS * -1) +
							"</font>%</p>";
					
					}

					if (c.MAX_DELTA_PROGRESS > -50 && c.MAX_DELTA_PROGRESS < -30)
					{
						delay += "<p> В период с " +
							c.MIN_PARCE_DATE.ToShortDateString() +
							" по " +
							c.MAX_PARCE_DATE.ToShortDateString() +
							" отставание по проекту<br>" +
							" №" +
							c.PROJ_REL_CAPTION + "<br>" +
							"<a href=\"ips://object/" + c.I_PROJ_ID.ToString() + "\">" + c.I_PROJ_CAPTION + "</a> " + //поменял на стринг							
							" увеличилось на " +
							Convert.ToString(((c.MAX_DELTA_PROGRESS) - (c.MIN_DELTA_PROGRESS)) * -1) +
							"%<br> " +
							"Отставание составляет: <font color=\"orangered\">" +
							Convert.ToString(c.MAX_DELTA_PROGRESS * -1) +
							"</font>%</p>";

						userdelay += "<p> В период с " +
							c.MIN_PARCE_DATE.ToShortDateString() +
							" по " +
							c.MAX_PARCE_DATE.ToShortDateString() +
							" отставание по проекту<br>" +
							" №" +
							c.PROJ_REL_CAPTION + "<br>" +
							"<a href=\"ips://object/" + c.I_PROJ_ID.ToString() + "\">" + c.I_PROJ_CAPTION + "</a> " + //поменял на стринг							
							" увеличилось на " +
							Convert.ToString(((c.MAX_DELTA_PROGRESS) - (c.MIN_DELTA_PROGRESS)) * -1) +
							"%<br> " +
							"Отставание составляет: <font color=\"orangered\">" +
							Convert.ToString(c.MAX_DELTA_PROGRESS * -1) +
							"</font>%</p>";
						
					}

					if (c.MAX_DELTA_PROGRESS < -50)
					{
						delay += "<p> В период с " +
							c.MIN_PARCE_DATE.ToShortDateString() +
							" по " +
							c.MAX_PARCE_DATE.ToShortDateString() +
							" отставание по проекту<br>" +
							" №" +
							c.PROJ_REL_CAPTION + "<br>" +
							"<a href=\"ips://object/" + c.I_PROJ_ID.ToString() + "\">" + c.I_PROJ_CAPTION + "</a> " + //поменял на стринг							
							" увеличилось на " +
							Convert.ToString(((c.MAX_DELTA_PROGRESS) - (c.MIN_DELTA_PROGRESS)) * -1) +
							"%<br> " +
							"Отставание составляет:  <font color=\"red\">" +
							Convert.ToString(c.MAX_DELTA_PROGRESS * -1) +
							"</font>%</p>";

						userdelay += "<p> В период с " +
							c.MIN_PARCE_DATE.ToShortDateString() +
							" по " +
							c.MAX_PARCE_DATE.ToShortDateString() +
							" отставание по проекту<br>" +
							" №" +
							c.PROJ_REL_CAPTION + "<br>" +
							"<a href=\"ips://object/" + c.I_PROJ_ID.ToString() + "\">" + c.I_PROJ_CAPTION + "</a> " + //поменял на стринг							
							" увеличилось на " +
							Convert.ToString(((c.MAX_DELTA_PROGRESS) - (c.MIN_DELTA_PROGRESS)) * -1) +
							"%<br> " +
							"Отставание составляет:  <font color=\"red\">" +
							Convert.ToString(c.MAX_DELTA_PROGRESS * -1) +
							"</font>%</p>";
						
					}
					if (c.MAX_DELTA_PROGRESS > 0)
					{
						delay += "<p> В период с " +
							c.MIN_PARCE_DATE.ToShortDateString() +
							" по " +
							c.MAX_PARCE_DATE.ToShortDateString() +
							" отставание по проекту<br>" +
							" №" +
							c.PROJ_REL_CAPTION + "<br>" +
							"<a href=\"ips://object/" + c.I_PROJ_ID.ToString() + "\">" + c.I_PROJ_CAPTION + "</a> " + //поменял на стринг							
							" сократилось на <font color=\"darkgreen\">" +
							Convert.ToString(((c.MAX_DELTA_PROGRESS) - (c.MIN_DELTA_PROGRESS)) * -1) +
							"</font>% <br>" + "Проект идет с опережением.";

						userdelay += "<p> В период с " +
							c.MIN_PARCE_DATE.ToShortDateString() +
							" по " +
							c.MAX_PARCE_DATE.ToShortDateString() +
							" отставание по проекту<br>" +
							" №" +
							c.PROJ_REL_CAPTION + "<br>" +
							"<a href=\"ips://object/" + c.I_PROJ_ID.ToString() + "\">" + c.I_PROJ_CAPTION + "</a> " + //поменял на стринг							
							" сократилось на <font color=\"darkgreen\">" +
							Convert.ToString(((c.MAX_DELTA_PROGRESS) - (c.MIN_DELTA_PROGRESS)) * -1) +
							"</font>% <br>" + "Проект идет с опережением.";
						
					}
				}

				if (c.MIN_DELTA_PROGRESS < 0 && c.MIN_DELTA_PROGRESS > -10)
				{
					if (c.MAX_DELTA_PROGRESS < -10 && c.MAX_DELTA_PROGRESS > -30)
					{
						delay += "<p> В период с " +
							c.MIN_PARCE_DATE.ToShortDateString() +
							" по " +
							c.MAX_PARCE_DATE.ToShortDateString() +
							" отставание по проекту<br>" +
							" №" +
							c.PROJ_REL_CAPTION + "<br>" +
							"<a href=\"ips://object/" + c.I_PROJ_ID.ToString() + "\">" + c.I_PROJ_CAPTION + "</a> " + //поменял на стринг							
							" увеличилось на " +
							Convert.ToString(((c.MAX_DELTA_PROGRESS) - (c.MIN_DELTA_PROGRESS)) * -1) +
							"%<br> " +
							"Отставание составляет:  <font color=\"sandybrown\">" +
							Convert.ToString(c.MAX_DELTA_PROGRESS * -1) +
							"</font>%</p>";

						userdelay += "<p> В период с " +
							c.MIN_PARCE_DATE.ToShortDateString() +
							" по " +
							c.MAX_PARCE_DATE.ToShortDateString() +
							" отставание по проекту<br>" +
							" №" +
							c.PROJ_REL_CAPTION + "<br>" +
							"<a href=\"ips://object/" + c.I_PROJ_ID.ToString() + "\">" + c.I_PROJ_CAPTION + "</a> " + //поменял на стринг							
							" увеличилось на " +
							Convert.ToString(((c.MAX_DELTA_PROGRESS) - (c.MIN_DELTA_PROGRESS)) * -1) +
							"%<br> " +
							"Отставание составляет:  <font color=\"sandybrown\">" +
							Convert.ToString(c.MAX_DELTA_PROGRESS * -1) +
							"</font>%</p>";
						
					}

					if (c.MAX_DELTA_PROGRESS < -30 && c.MAX_DELTA_PROGRESS > -50)
					{
						delay += "<p> В период с " +
							c.MIN_PARCE_DATE.ToShortDateString() +
							" по " +
							c.MAX_PARCE_DATE.ToShortDateString() +
							" отставание по проекту<br>" +
							" №" +
							c.PROJ_REL_CAPTION + "<br>" +
							"<a href=\"ips://object/" + c.I_PROJ_ID.ToString() + "\">" + c.I_PROJ_CAPTION + "</a> " + //поменял на стринг							
							" увеличилось на " +
							Convert.ToString(((c.MAX_DELTA_PROGRESS) - (c.MIN_DELTA_PROGRESS)) * -1) +
							"%<br> " +
							"Отставание составляет:  <font color=\"orangered\">" +
							Convert.ToString(c.MAX_DELTA_PROGRESS * -1) +
							"</font>%</p>";

						userdelay += "<p> В период с " +
							c.MIN_PARCE_DATE.ToShortDateString() +
							" по " +
							c.MAX_PARCE_DATE.ToShortDateString() +
							" отставание по проекту<br>" +
							" №" +
							c.PROJ_REL_CAPTION + "<br>" +
							"<a href=\"ips://object/" + c.I_PROJ_ID.ToString() + "\">" + c.I_PROJ_CAPTION + "</a> " + //поменял на стринг							
							" увеличилось на " +
							Convert.ToString(((c.MAX_DELTA_PROGRESS) - (c.MIN_DELTA_PROGRESS)) * -1) +
							"%<br> " +
							"Отставание составляет:  <font color=\"orangered\">" +
							Convert.ToString(c.MAX_DELTA_PROGRESS * -1) +
							"</font>%</p>";
						
					}
					if (c.MAX_DELTA_PROGRESS < -50)
					{
						delay += "<p> В период с " +
							c.MIN_PARCE_DATE.ToShortDateString() +
							" по " +
							c.MAX_PARCE_DATE.ToShortDateString() +
							" отставание по проекту<br>" +
							" №" +
							c.PROJ_REL_CAPTION + "<br>" +
							"<a href=\"ips://object/" + c.I_PROJ_ID.ToString() + "\">" + c.I_PROJ_CAPTION + "</a> " + //поменял на стринг							
							" увеличилось на " +
							Convert.ToString(((c.MAX_DELTA_PROGRESS) - (c.MIN_DELTA_PROGRESS)) * -1) +
							"%<br> " +
							"Отставание составляет: <font color=\"red\">" +
							Convert.ToString(c.MAX_DELTA_PROGRESS * -1) +
							"</font>%</p>";

						userdelay += "<p> В период с " +
							c.MIN_PARCE_DATE.ToShortDateString() +
							" по " +
							c.MAX_PARCE_DATE.ToShortDateString() +
							" отставание по проекту<br>" +
							" №" +
							c.PROJ_REL_CAPTION + "<br>" +
							"<a href=\"ips://object/" + c.I_PROJ_ID.ToString() + "\">" + c.I_PROJ_CAPTION + "</a> " + //поменял на стринг							
							" увеличилось на " +
							Convert.ToString(((c.MAX_DELTA_PROGRESS) - (c.MIN_DELTA_PROGRESS)) * -1) +
							"%<br> " +
							"Отставание составляет: <font color=\"red\">" +
							Convert.ToString(c.MAX_DELTA_PROGRESS * -1) +
							"</font>%</p>";
						
					}
	
					if (c.MAX_DELTA_PROGRESS > 0)
					{
						delay += "<p> В период с " +
							c.MIN_PARCE_DATE.ToShortDateString() +
							" по " +
							c.MAX_PARCE_DATE.ToShortDateString() +
							" отставание по проекту<br>" +
							" №" +
							c.PROJ_REL_CAPTION + "<br>" +
							"<a href=\"ips://object/" + c.I_PROJ_ID.ToString() + "\">" + c.I_PROJ_CAPTION + "</a> " + //поменял на стринг							
							" сократилось на <font color=\"darkgreen\">" +
							Convert.ToString(((c.MAX_DELTA_PROGRESS) - (c.MIN_DELTA_PROGRESS)) * -1) +
							"</font>% <br>" + "Проект идет с опережением.";
							
						userdelay += "<p> В период с " +
							c.MIN_PARCE_DATE.ToShortDateString() +
							" по " +
							c.MAX_PARCE_DATE.ToShortDateString() +
							" отставание по проекту<br>" +
							" №" +
							c.PROJ_REL_CAPTION + "<br>" +
							"<a href=\"ips://object/" + c.I_PROJ_ID.ToString() + "\">" + c.I_PROJ_CAPTION + "</a> " + //поменял на стринг							
							" сократилось на <font color=\"darkgreen\">" +
							Convert.ToString(((c.MAX_DELTA_PROGRESS) - (c.MIN_DELTA_PROGRESS)) * -1) +
							"</font>% <br>" + "Проект идет с опережением.";
						
					}
					
				}
				
				if (c.MIN_DELTA_PROGRESS < -10 && c.MIN_DELTA_PROGRESS > -30)
				{

					if (c.MAX_DELTA_PROGRESS < -30 && c.MAX_DELTA_PROGRESS > -50)
					{
						delay += "<p> В период с " +
							c.MIN_PARCE_DATE.ToShortDateString() +
							" по " +
							c.MAX_PARCE_DATE.ToShortDateString() +
							" отставание по проекту<br>" +
							" №" +
							c.PROJ_REL_CAPTION + "<br>" +
							"<a href=\"ips://object/" + c.I_PROJ_ID.ToString() + "\">" + c.I_PROJ_CAPTION + "</a> " + //поменял на стринг							
							" увеличилось на " +
							Convert.ToString(((c.MAX_DELTA_PROGRESS) - (c.MIN_DELTA_PROGRESS)) * -1) +
							"% <br>" +
							"Отставание составляет: <font color=\"darkorange\">" +
							Convert.ToString(c.MAX_DELTA_PROGRESS * -1) +
							"</font>%</p>";

						userdelay += "<p> В период с " +
							c.MIN_PARCE_DATE.ToShortDateString() +
							" по " +
							c.MAX_PARCE_DATE.ToShortDateString() +
							" отставание по проекту<br>" +
							" №" +
							c.PROJ_REL_CAPTION + "<br>" +
							"<a href=\"ips://object/" + c.I_PROJ_ID.ToString() + "\">" + c.I_PROJ_CAPTION + "</a> " + //поменял на стринг							
							" увеличилось на " +
							Convert.ToString(((c.MAX_DELTA_PROGRESS) - (c.MIN_DELTA_PROGRESS)) * -1) +
							"% <br>" +
							"Отставание составляет: <font color=\"darkorange\">" +
							Convert.ToString(c.MAX_DELTA_PROGRESS * -1) +
							"</font>%</p>";
						
					}

					
					if (c.MAX_DELTA_PROGRESS < -50)
					{
						delay += "<p> В период с " +
							c.MIN_PARCE_DATE.ToShortDateString() +
							" по " +
							c.MAX_PARCE_DATE.ToShortDateString() +
							" отставание по проекту<br>" +
							" №" +
							c.PROJ_REL_CAPTION + "<br>" +
							"<a href=\"ips://object/" + c.I_PROJ_ID.ToString() + "\">" + c.I_PROJ_CAPTION + "</a> " + //поменял на стринг							
							" увеличилось на " +
							Convert.ToString(((c.MAX_DELTA_PROGRESS) - (c.MIN_DELTA_PROGRESS)) * -1) +
							"% <br>" +
							"Отставание составляет: <font color=\"red\">" +
							Convert.ToString(c.MAX_DELTA_PROGRESS * -1) +
							"</font> %</p>";

						userdelay += "<p> В период с " +
							c.MIN_PARCE_DATE.ToShortDateString() +
							" по " +
							c.MAX_PARCE_DATE.ToShortDateString() +
							" отставание по проекту<br>" +
							" №" +
							c.PROJ_REL_CAPTION + "<br>" +
							"<a href=\"ips://object/" + c.I_PROJ_ID.ToString() + "\">" + c.I_PROJ_CAPTION + "</a> " + //поменял на стринг							
							" увеличилось на " +
							Convert.ToString(((c.MAX_DELTA_PROGRESS) - (c.MIN_DELTA_PROGRESS)) * -1) +
							"% <br>" +
							"Отставание составляет: <font color=\"red\">" +
							Convert.ToString(c.MAX_DELTA_PROGRESS * -1) +
							"</font> %</p>";
						
					}
					 
					
				}

				if (c.MIN_DELTA_PROGRESS < -30 && c.MIN_DELTA_PROGRESS > -50)
				{
					if (c.MAX_DELTA_PROGRESS < -50)
					{
						delay += "<p> В период с " +
							c.MIN_PARCE_DATE.ToShortDateString() +
							" по " +
							c.MAX_PARCE_DATE.ToShortDateString() +
							" отставание по проекту<br>" +
							" №" +
							c.PROJ_REL_CAPTION + "<br>" +
							"<a href=\"ips://object/" + c.I_PROJ_ID.ToString() + "\">" + c.I_PROJ_CAPTION + "</a> " + //поменял на стринг							
							" увеличилось на " +
							Convert.ToString(((c.MAX_DELTA_PROGRESS) - (c.MIN_DELTA_PROGRESS)) * -1) +
							"%<br> " +
							"Отставание составляет: <font color=\"red\">" +
							Convert.ToString(c.MAX_DELTA_PROGRESS * -1) +
							"</font>%</p>";

						userdelay += "<p> В период с " +
							c.MIN_PARCE_DATE.ToShortDateString() +
							" по " +
							c.MAX_PARCE_DATE.ToShortDateString() +
							" отставание по проекту<br>" +
							" №" +
							c.PROJ_REL_CAPTION + "<br>" +
							"<a href=\"ips://object/" + c.I_PROJ_ID.ToString() + "\">" + c.I_PROJ_CAPTION + "</a> " + //поменял на стринг							
							" увеличилось на " +
							Convert.ToString(((c.MAX_DELTA_PROGRESS) - (c.MIN_DELTA_PROGRESS)) * -1) +
							"%<br> " +
							"Отставание составляет: <font color=\"red\">" +
							Convert.ToString(c.MAX_DELTA_PROGRESS * -1) +
							"</font>%</p>";
						
					}


				}

				//Оповещение об отставании проекта более 50%

				if (c.MAX_DELTA_PROGRESS < -50)
				{

					if (maxdelay == String.Empty)
					{
						maxdelay = "<p><b>По данным ресурсно-календарного плана IPS у вас есть критическое отставание по проектам свыше 50%!:</b><br><b>" + 
							num.ToString() +
							".</b>  " + 
							"№ договора : " + 
							c.PROJ_REL_CAPTION +
							"<br>Ресурсно-календарный план : <br>" +
							"<a href=\"ips://object/" + c.I_PROJ_ID.ToString() + "\">" + c.I_PROJ_CAPTION + "</a><br> " + //поменял на стринг	
							"Отставание:  <font color=\"red\">" + 
							c.MAX_DELTA_PROGRESS * (-1) + 
							"</font>%</p>";

						usermaxdelay = "<p><b>По данным ресурсно-календарного плана IPS у вас есть критическое отставание по проектам свыше 50%!:</b><br><b>" +
							num.ToString() +
							".</b>  " +
							"№ договора : " +
							c.PROJ_REL_CAPTION +
							"<br>Ресурсно-календарный план : <br>" +
							"<a href=\"ips://object/" + c.I_PROJ_ID.ToString() + "\">" + c.I_PROJ_CAPTION + "</a><br> " + //поменял на стринг	
							"Отставание:  <font color=\"red\">" +
							c.MAX_DELTA_PROGRESS * (-1) +
							"</font>%</p>";

						i++;
						num++;
					}
					else
					{
						maxdelay += "<p><b>" +
							num.ToString() +
							".</b>  " +
							"№ договора : " +
							c.PROJ_REL_CAPTION +
							"<br>Ресурсно-календарный план : <br>" +
							"<a href=\"ips://object/" + c.I_PROJ_ID.ToString() + "\">" + c.I_PROJ_CAPTION + "</a><br> " + //поменял на стринг	
							"Отставание:  <font color=\"red\">" +
							c.MAX_DELTA_PROGRESS * (-1) +
							"</font>%</p>";

						usermaxdelay += "<p><b>" +
							num.ToString() +
							".</b>  " +
							"№ договора : " +
							c.PROJ_REL_CAPTION +
							"<br>Ресурсно-календарный план : <br>" +
							"<a href=\"ips://object/" + c.I_PROJ_ID.ToString() + "\">" + c.I_PROJ_CAPTION + "</a><br> " + //поменял на стринг	
							"Отставание:  <font color=\"red\">" +
							c.MAX_DELTA_PROGRESS * (-1) +
							"</font>%</p>";

						i++;
						num++;
					}

				}

				if (c.I_FACT_PROGRESS < 100 && c.I_PLAN_FINISH < nowdate)
				{
					if (overdue == String.Empty)
					{
						overdue = "<p><b>У Вас есть проекты с просроченной плановой датой завершения по РКП:<br></b><b>" + 
							numoverdue.ToString() + 
							".</b>  " + "№ договора :" +
							c.PROJ_REL_CAPTION + 
							"<br>Ресурсно-календарный план :<br> " + 
							"<a href=\"ips://object/" + c.I_PROJ_ID.ToString() + "\">" + c.I_PROJ_CAPTION + "</a><br> " +
							"Дата завершения РКП: <font color=\"red\">" + 
							c.I_PLAN_FINISH.ToShortDateString() + 
							"</font></p>";

						useroverdue = "<p><b>У Вас есть проекты с просроченной плановой датой завершения по РКП:<br></b><b>" +
							numoverdue.ToString() +
							".</b>  " + "№ договора :" +
							c.PROJ_REL_CAPTION +
							"<br>Ресурсно-календарный план :<br> " +
							"<a href=\"ips://object/" + c.I_PROJ_ID.ToString() + "\">" + c.I_PROJ_CAPTION + "</a><br> " +
							"Дата завершения РКП: <font color=\"red\">" +
							c.I_PLAN_FINISH.ToShortDateString() +
							"</font></p>";

						numoverdue++;
					}
					else
					{
						overdue += "<p><b>" + 
							numoverdue.ToString() + 
							".</b>  " + "№ договора :" + 
							c.PROJ_REL_CAPTION +
							"<br>Ресурсно-календарный план :<br> " +
							"<a href=\"ips://object/" + c.I_PROJ_ID.ToString() + "\">" + c.I_PROJ_CAPTION + "</a><br> " +
							"Дата завершения РКП: <font color=\"red\">" +
							c.I_PLAN_FINISH.ToShortDateString() +
							"</font></p>";

						useroverdue += "<p><b>" +
							numoverdue.ToString() +
							".</b>  " + "№ договора :" +
							c.PROJ_REL_CAPTION +
							"<br>Ресурсно-календарный план :<br> " +
							"<a href=\"ips://object/" + c.I_PROJ_ID.ToString() + "\">" + c.I_PROJ_CAPTION + "</a><br> " +
							"Дата завершения РКП: <font color=\"red\">" +
							c.I_PLAN_FINISH.ToShortDateString() +
							"</font></p>";

						numoverdue++;
					}

				}			
			}

			if (delay == String.Empty && maxdelay == String.Empty && overdue == String.Empty)
				continue;

			if (log != String.Empty)
			{
				log += "<p>" + mailadress + "</p>" + captionname + "\n" + delay;
				if (maxdelay != String.Empty)
				{
					log += maxdelay;
				}
				if (overdue != String.Empty)
				{
					log += overdue;
				}

			}
			else
			{
				log += "<p>" + mailadress + "</p>" + captionname + "\n" + delay;
				if (maxdelay != String.Empty)
				{
					log += maxdelay;
				}
				if (overdue != String.Empty)
				{
					log += overdue;
				}

			}

			//Отправка админу динамики выполнения проектов
			if (delay != String.Empty)
			{

					dynamiclog += "<p>" + nameforadmin + "</p>" + delay;
 
			}

		

			string userlog = string.Empty;
			userlog = captionname + userdelay + usermaxdelay + useroverdue;
	
			
			SmtpClient SmtptoUser = new SmtpClient("mailserver");
			
			MailMessage MessagetoUser = new MailMessage();			
			MessagetoUser.From = new MailAddress("mailaddress@mail.com");
			MessagetoUser.To.Add(new MailAddress(mailadress));		
			MessagetoUser.IsBodyHtml = true;
			 if (userlog != String.Empty)
			{
				SmtptoUser.Send(MessagetoUser);
				MessagetoUser.Dispose();
			log += "<br>" + "<p><b><font color=\"darkgreen\">Отправлено! "+"на адрес : " + mailadress + "</font></b></p>";
			}

			
		}


		SmtpClient SmtptoAdmin = new SmtpClient("mailserver");
		
		MailMessage MessagetoAdmin = new MailMessage();
        MessagetoAdmin.From = new MailAddress("mailaddress@mail.com");
        MessagetoAdmin.To.Add(new MailAddress("mailaddress@mail.com"));
		
		
		MessagetoAdmin.IsBodyHtml = true;
		if (log != String.Empty)
		{
			SmtptoAdmin.Send(MessagetoAdmin);
		
		}

	}

}


namespace CsSQL
{
	public class DBSQLServerUtils
	{
		public static SqlConnection GetDBConnection(string datasource, string database, string username, string password)
		{
			
			string connString = @"Data Source=" + datasource + ";Initial Catalog=" + database + ";Persist Security Info= no ;User ID = " + username + "; Password = " + password;
			SqlConnection conn = new SqlConnection(connString);
			//Console.WriteLine(connString);
			return conn;
		}
	}
	public class DBUtils
	{
		public static SqlConnection GetDBConnection()
		{
			string datasource = @"TA03";
			string database = "t_IMProject";
			string username = "ea";
			//Console.WriteLine(username);
			string password = "8fuckofF8";
			return DBSQLServerUtils.GetDBConnection(datasource, database, username, password);
		}
	}
	
}


namespace L2SApp
{

	[Table(Name = "PersentForWeek")]
	public class PersentForWeeks
	{

		[Column]
		public double I_PROJ_ID { get; set; }
		[Column]
		public string PROJ_REL_CAPTION { get; set; }
		[Column]
		public string I_PROJ_CAPTION { get; set; }
		[Column]
		public decimal? I_PLAN_PROGRESS { get; set; }
		[Column]
		public decimal? I_FACT_PROGRESS { get; set; }
		[Column]
		public DateTime I_PLAN_FINISH { get; set; }
		[Column]
		public string USER_NAME { get; set; }
		[Column]
		public double USER_ID { get; set; }
		[Column]
		public string USER_MAIL { get; set; }
		[Column]
		public DateTime MIN_PARCE_DATE { get; set; }
		[Column]
		public decimal? MIN_DELTA_PROGRESS { get; set; }
		[Column]
		public DateTime MAX_PARCE_DATE { get; set; }
		[Column]
		public decimal? MAX_DELTA_PROGRESS { get; set; }


	}

}