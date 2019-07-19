//============================================//
//数据库操作类
//HYUO
//2014-08-30
//Ver 1.0.1.2356
//============================================//
using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Text;

namespace QianSheng_Data_Export.Common.CommDB
{
	/// <summary>
	/// 数据库操作类
	/// </summary>
	public class DbOperation
	{
		//从配置文件(App.config)中获取连接字符串
		public static string connectionString = ConfigurationManager.AppSettings["CASCRS_DB_CONN"].ToString();
	
		//定义SqlConnection
		private static SqlConnection objSqlConn = null;
		//定义SqlDataAdapter
		private static SqlDataAdapter objSqlDataAdapter = null;
		//定义SqlCommandBuilder
		private static SqlCommandBuilder objSqlCommandBuilder = null;

		/// <summary>
		/// 连接数据库
		/// </summary>
		private static void GetConnection()
		{
			try
			{
				objSqlConn = new SqlConnection(connectionString);
				objSqlConn.Open();
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		

		/// <summary>
		/// 根据查询SQL字符串，获取DataTable
		/// </summary>
		///<param name = "strSelectSql">查询字符串</param>
		///<returns>返回DataSet</returns>
		public static DataTable GetDataTable(string strSelectSql, int ConnId = 0, SqlConnection objExtConn = null)
		{
			try
			{
				DataTable dtTarget = new DataTable();
				if (objExtConn == null)
				{
					if (ConnId == 0)
						GetConnection();
					//else if (ConnId == 1)
					//	GetConnection1();
					//else
					//	GetConnection2();
					objSqlDataAdapter = new SqlDataAdapter(strSelectSql, objSqlConn);
					objSqlDataAdapter.Fill(dtTarget);
				}
				else
				{
					if (objExtConn.State == ConnectionState.Closed)
						objExtConn.Open();
					SqlCommand objSqlCmd = new SqlCommand(strSelectSql, objExtConn);
					IAsyncResult result = objSqlCmd.BeginExecuteReader(new AsyncCallback((ar) => { }), objSqlCmd);
					SqlDataReader sdr = objSqlCmd.EndExecuteReader(result);
					dtTarget.Load(sdr);
				}
				return dtTarget;
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		/// <summary>
		/// 根据查询SQL字符串，获取DataTable
		/// </summary>
		/// ///<param name = "objSqlDataAdapter">数据适配器</param>
		///<param name = "strSelectSql">查询字符串</param>
		///<returns>返回DataSet</returns>
		public static DataTable GetDataTable(out SqlDataAdapter objSqlDataAdapter, string strSelectSql, int ConnId = 0)
		{
			try
			{
				if (ConnId == 0)
					GetConnection();
				
				DataTable dtTarget = new DataTable();
				objSqlDataAdapter = new SqlDataAdapter(strSelectSql, objSqlConn);
				objSqlDataAdapter.Fill(dtTarget);
				return dtTarget;
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		/// <summary>
		/// 将DataGridView的数据源更新到数据库中
		/// </summary>
		/// <param name="dtResult">数据源</param>
		/// <returns>执行状态</returns>
		public static void UpdateDataSource(SqlDataAdapter objSqlDataAdapter, DataTable dtResult)
		{
			try
			{
				objSqlDataAdapter.ContinueUpdateOnError = true;
				using (objSqlCommandBuilder = new SqlCommandBuilder(objSqlDataAdapter))
				{
					objSqlDataAdapter.DeleteCommand = objSqlCommandBuilder.GetDeleteCommand();
					objSqlDataAdapter.InsertCommand = objSqlCommandBuilder.GetInsertCommand();
					objSqlDataAdapter.UpdateCommand = objSqlCommandBuilder.GetUpdateCommand();
					objSqlDataAdapter.Update(dtResult);
					dtResult.AcceptChanges();
				}
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		/// <summary>
		/// 执行非查询SQL字符串，如Insert，Update，Delete
		/// </summary>
		///<param name = "strSql">非查询SQL字符串</param>
		///<returns>返回受影响的行数</returns>
		public static int ExecuteNonQuery(string strSql, int ConnId = 0, SqlConnection objExtConn = null)
		{
			try
			{
				int intResult = 0;
				if (objExtConn == null)
				{
					if (ConnId == 0)
						GetConnection();
					
					using (SqlCommand objSqlCmd = new SqlCommand(strSql, objSqlConn))
					{
						objSqlCmd.CommandTimeout = 3600;
						objSqlCmd.ExecuteNonQuery();
					}
				}
				else
				{
					if (objExtConn.State == ConnectionState.Closed)
						objExtConn.Open();
					using (SqlCommand objSqlCmd = new SqlCommand(strSql, objExtConn))
					{
						objSqlCmd.ExecuteNonQuery();
					}
				}

				return intResult;
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		/// <summary>
		/// 执行数据快速插入或复制,要求数据源DataTable的表结构与目标表完全一致
		/// </summary>
		///<param name = "dtSource">数据源DataTable</param>
		///<param name = "strDestinationTableName">目标表名称</param>
		///<returns>返回执行状态</returns>
		public static bool ExecuteSqlBulkCopy(DataTable dtSource, string strDestinationTableName, int ConnId = 0)
		{
			try
			{
				if (ConnId == 0)
					GetConnection();
			
				using (SqlBulkCopy objSqlBulkCopy = new SqlBulkCopy(objSqlConn) {  })
				{
					objSqlBulkCopy.DestinationTableName = strDestinationTableName;
					for (int i = 0; i < dtSource.Columns.Count; i++)
					{
						objSqlBulkCopy.ColumnMappings.Add(i, i);
					}
					objSqlBulkCopy.WriteToServer(dtSource);
				}
				return true;
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		/// <summary>
		/// 执行存储过程
		/// </summary>
		/// <param name="strStoreProcName">存储过程名称</param>
		/// <param name="strParametersName">参数名集合</param>
		/// <param name="strParameterValue">参数值集合</param>
		/// <returns>执行状态</returns>
		public static bool ExecuteStoreProcedure(string strStoreProcName, string[] strParametersName, string[] strParameterValue)
		{
			try
			{
				GetConnection();
				using (SqlCommand objSqlCmd = new SqlCommand(strStoreProcName, objSqlConn))
				{
					objSqlCmd.CommandType = CommandType.StoredProcedure;
					int intParaCount = strParametersName.Length;
					for (int i = 0; i < intParaCount; i++)
					{
						objSqlCmd.Parameters.Add(strParametersName[i], SqlDbType.NVarChar);
						objSqlCmd.Parameters[strParametersName[i]].Value = strParameterValue[i];
					}
					objSqlCmd.ExecuteNonQuery();
				}
				return true;
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		/// <summary>
		/// 根据传过来的数据源的数据结构创建临时表，并将数据源中的数据写入到创建的临时表中
		/// </summary>
		/// <param name="dtSource">数据源</param>
		/// <param name="TableName">临时表名</param>
		public static void DataTableToDB(DataTable dtSource, string TableName)
		{
			try
			{
				StringBuilder sbSQL = new StringBuilder();
				sbSQL.AppendLine(string.Format("CREATE TABLE [dbo].[{0}] (", TableName));
				foreach (DataColumn dc in dtSource.Columns)
				{
					if (Type.GetType("System.String").Equals(dc.DataType))
						sbSQL.AppendLine(string.Format("    [{0}] NVARCHAR(50) NULL,", dc.ColumnName));
					else if (Type.GetType("System.Int32").Equals(dc.DataType) || Type.GetType("System.Int64").Equals(dc.DataType))
						sbSQL.AppendLine(string.Format("    [{0}] INT NULL,", dc.ColumnName));
					else if (Type.GetType("System.Double").Equals(dc.DataType))
						sbSQL.AppendLine(string.Format("    [{0}] FLOAT NULL,", dc.ColumnName));
					else if (Type.GetType("System.Decimal").Equals(dc.DataType))
						sbSQL.AppendLine(string.Format("    [{0}] DECIMAL(18, 2) NULL,", dc.ColumnName));
				}
				sbSQL.AppendLine(")");
				string strSQL = sbSQL.ToString();
				ExecuteNonQuery(strSQL);
				ExecuteSqlBulkCopy(dtSource, TableName);
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		/// <summary>
		/// 执行非查询SQL脚本文件
		/// </summary>
		/// <param name="FileName">脚本文件路径</param>
		public static void ExecuteNonQuerySqlScript(string FileName, string DbName, string ObjectName, int ConnId = 0)
		{
			try
			{
				if (ConnId == 0)
					GetConnection();
				
				//从SQL文件获取SQL语句
				StreamReader sr = new StreamReader(FileName, Encoding.Default);
				string SqlScript = sr.ReadToEnd();
				if (DbName != null && ObjectName != null)
				{
					SqlScript = SqlScript.Replace(string.Format("[dbo].[{0}]", ObjectName), string.Format("[{0}].[dbo].[{1}]", DbName, ObjectName));
				}
				//执行SQL语句
				ExecuteNonQuery(SqlScript);
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}
	}
}