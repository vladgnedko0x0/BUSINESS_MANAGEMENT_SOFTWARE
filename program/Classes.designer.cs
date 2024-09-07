﻿#pragma warning disable 1591
//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace program
{
	using System.Data.Linq;
	using System.Data.Linq.Mapping;
	using System.Data;
	using System.Collections.Generic;
	using System.Reflection;
	using System.Linq;
	using System.Linq.Expressions;
	using System.ComponentModel;
	using System;
	
	
	[global::System.Data.Linq.Mapping.DatabaseAttribute(Name="Firma")]
	public partial class ClassesDataContext : System.Data.Linq.DataContext
	{
		
		private static System.Data.Linq.Mapping.MappingSource mappingSource = new AttributeMappingSource();
		
    #region Extensibility Method Definitions
    partial void OnCreated();
    partial void InsertItem(Item instance);
    partial void UpdateItem(Item instance);
    partial void DeleteItem(Item instance);
    partial void InsertCompany(Company instance);
    partial void UpdateCompany(Company instance);
    partial void DeleteCompany(Company instance);
    partial void InsertInnerOrder(InnerOrder instance);
    partial void UpdateInnerOrder(InnerOrder instance);
    partial void DeleteInnerOrder(InnerOrder instance);
    partial void InsertItemOrder(ItemOrder instance);
    partial void UpdateItemOrder(ItemOrder instance);
    partial void DeleteItemOrder(ItemOrder instance);
    #endregion
		
		public ClassesDataContext() : 
				base(global::program.Properties.Settings.Default.FirmaConnectionString1, mappingSource)
		{
			OnCreated();
		}
		
		public ClassesDataContext(string connection) : 
				base(connection, mappingSource)
		{
			OnCreated();
		}
		
		public ClassesDataContext(System.Data.IDbConnection connection) : 
				base(connection, mappingSource)
		{
			OnCreated();
		}
		
		public ClassesDataContext(string connection, System.Data.Linq.Mapping.MappingSource mappingSource) : 
				base(connection, mappingSource)
		{
			OnCreated();
		}
		
		public ClassesDataContext(System.Data.IDbConnection connection, System.Data.Linq.Mapping.MappingSource mappingSource) : 
				base(connection, mappingSource)
		{
			OnCreated();
		}
		
		public System.Data.Linq.Table<Item> Item
		{
			get
			{
				return this.GetTable<Item>();
			}
		}
		
		public System.Data.Linq.Table<Company> Companies
		{
			get
			{
				return this.GetTable<Company>();
			}
		}
		
		public System.Data.Linq.Table<InnerOrder> InnerOrders
		{
			get
			{
				return this.GetTable<InnerOrder>();
			}
		}
		
		public System.Data.Linq.Table<ItemOrder> ItemOrders
		{
			get
			{
				return this.GetTable<ItemOrder>();
			}
		}
	}
	
	[global::System.Data.Linq.Mapping.TableAttribute(Name="dbo.Item")]
	public partial class Item : INotifyPropertyChanging, INotifyPropertyChanged
	{
		
		private static PropertyChangingEventArgs emptyChangingEventArgs = new PropertyChangingEventArgs(String.Empty);
		
		private int _Id;
		
		private string _ItemCode;
		
		private double _Size;
		
		private string _Image;
		
		private EntitySet<InnerOrder> _InnerOrders;
		
		private EntitySet<ItemOrder> _ItemOrders;
		
    #region Extensibility Method Definitions
    partial void OnLoaded();
    partial void OnValidate(System.Data.Linq.ChangeAction action);
    partial void OnCreated();
    partial void OnIdChanging(int value);
    partial void OnIdChanged();
    partial void OnItemCodeChanging(string value);
    partial void OnItemCodeChanged();
    partial void OnSizeChanging(double value);
    partial void OnSizeChanged();
    partial void OnImageChanging(string value);
    partial void OnImageChanged();
    #endregion
		
		public Item()
		{
			this._InnerOrders = new EntitySet<InnerOrder>(new Action<InnerOrder>(this.attach_InnerOrders), new Action<InnerOrder>(this.detach_InnerOrders));
			this._ItemOrders = new EntitySet<ItemOrder>(new Action<ItemOrder>(this.attach_ItemOrders), new Action<ItemOrder>(this.detach_ItemOrders));
			OnCreated();
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_Id", AutoSync=AutoSync.OnInsert, DbType="Int NOT NULL IDENTITY", IsPrimaryKey=true, IsDbGenerated=true)]
		public int Id
		{
			get
			{
				return this._Id;
			}
			set
			{
				if ((this._Id != value))
				{
					this.OnIdChanging(value);
					this.SendPropertyChanging();
					this._Id = value;
					this.SendPropertyChanged("Id");
					this.OnIdChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_ItemCode", DbType="NVarChar(MAX) NOT NULL", CanBeNull=false)]
		public string ItemCode
		{
			get
			{
				return this._ItemCode;
			}
			set
			{
				if ((this._ItemCode != value))
				{
					this.OnItemCodeChanging(value);
					this.SendPropertyChanging();
					this._ItemCode = value;
					this.SendPropertyChanged("ItemCode");
					this.OnItemCodeChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_Size", DbType="Float NOT NULL")]
		public double Size
		{
			get
			{
				return this._Size;
			}
			set
			{
				if ((this._Size != value))
				{
					this.OnSizeChanging(value);
					this.SendPropertyChanging();
					this._Size = value;
					this.SendPropertyChanged("Size");
					this.OnSizeChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_Image", DbType="NVarChar(MAX)")]
		public string Image
		{
			get
			{
				return this._Image;
			}
			set
			{
				if ((this._Image != value))
				{
					this.OnImageChanging(value);
					this.SendPropertyChanging();
					this._Image = value;
					this.SendPropertyChanged("Image");
					this.OnImageChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.AssociationAttribute(Name="Item_InnerOrder", Storage="_InnerOrders", ThisKey="Id", OtherKey="ItemId")]
		public EntitySet<InnerOrder> InnerOrders
		{
			get
			{
				return this._InnerOrders;
			}
			set
			{
				this._InnerOrders.Assign(value);
			}
		}
		
		[global::System.Data.Linq.Mapping.AssociationAttribute(Name="Item_ItemOrder", Storage="_ItemOrders", ThisKey="Id", OtherKey="ItemId")]
		public EntitySet<ItemOrder> ItemOrders
		{
			get
			{
				return this._ItemOrders;
			}
			set
			{
				this._ItemOrders.Assign(value);
			}
		}
		
		public event PropertyChangingEventHandler PropertyChanging;
		
		public event PropertyChangedEventHandler PropertyChanged;
		
		protected virtual void SendPropertyChanging()
		{
			if ((this.PropertyChanging != null))
			{
				this.PropertyChanging(this, emptyChangingEventArgs);
			}
		}
		
		protected virtual void SendPropertyChanged(String propertyName)
		{
			if ((this.PropertyChanged != null))
			{
				this.PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
			}
		}
		
		private void attach_InnerOrders(InnerOrder entity)
		{
			this.SendPropertyChanging();
			entity.Item = this;
		}
		
		private void detach_InnerOrders(InnerOrder entity)
		{
			this.SendPropertyChanging();
			entity.Item = null;
		}
		
		private void attach_ItemOrders(ItemOrder entity)
		{
			this.SendPropertyChanging();
			entity.Item = this;
		}
		
		private void detach_ItemOrders(ItemOrder entity)
		{
			this.SendPropertyChanging();
			entity.Item = null;
		}
	}
	
	[global::System.Data.Linq.Mapping.TableAttribute(Name="dbo.Companies")]
	public partial class Company : INotifyPropertyChanging, INotifyPropertyChanged
	{
		
		private static PropertyChangingEventArgs emptyChangingEventArgs = new PropertyChangingEventArgs(String.Empty);
		
		private int _Id;
		
		private string _Name;
		
		private string _ShortName;
		
		private string _Addres;
		
		private System.Nullable<int> _NIP;
		
		private EntitySet<ItemOrder> _ItemOrders;
		
    #region Extensibility Method Definitions
    partial void OnLoaded();
    partial void OnValidate(System.Data.Linq.ChangeAction action);
    partial void OnCreated();
    partial void OnIdChanging(int value);
    partial void OnIdChanged();
    partial void OnNameChanging(string value);
    partial void OnNameChanged();
    partial void OnShortNameChanging(string value);
    partial void OnShortNameChanged();
    partial void OnAddresChanging(string value);
    partial void OnAddresChanged();
    partial void OnNIPChanging(System.Nullable<int> value);
    partial void OnNIPChanged();
    #endregion
		
		public Company()
		{
			this._ItemOrders = new EntitySet<ItemOrder>(new Action<ItemOrder>(this.attach_ItemOrders), new Action<ItemOrder>(this.detach_ItemOrders));
			OnCreated();
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_Id", AutoSync=AutoSync.OnInsert, DbType="Int NOT NULL IDENTITY", IsPrimaryKey=true, IsDbGenerated=true)]
		public int Id
		{
			get
			{
				return this._Id;
			}
			set
			{
				if ((this._Id != value))
				{
					this.OnIdChanging(value);
					this.SendPropertyChanging();
					this._Id = value;
					this.SendPropertyChanged("Id");
					this.OnIdChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_Name", DbType="NVarChar(MAX) NOT NULL", CanBeNull=false)]
		public string Name
		{
			get
			{
				return this._Name;
			}
			set
			{
				if ((this._Name != value))
				{
					this.OnNameChanging(value);
					this.SendPropertyChanging();
					this._Name = value;
					this.SendPropertyChanged("Name");
					this.OnNameChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_ShortName", DbType="NVarChar(MAX) NOT NULL", CanBeNull=false)]
		public string ShortName
		{
			get
			{
				return this._ShortName;
			}
			set
			{
				if ((this._ShortName != value))
				{
					this.OnShortNameChanging(value);
					this.SendPropertyChanging();
					this._ShortName = value;
					this.SendPropertyChanged("ShortName");
					this.OnShortNameChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_Addres", DbType="NVarChar(MAX) NOT NULL", CanBeNull=false)]
		public string Addres
		{
			get
			{
				return this._Addres;
			}
			set
			{
				if ((this._Addres != value))
				{
					this.OnAddresChanging(value);
					this.SendPropertyChanging();
					this._Addres = value;
					this.SendPropertyChanged("Addres");
					this.OnAddresChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_NIP", DbType="Int")]
		public System.Nullable<int> NIP
		{
			get
			{
				return this._NIP;
			}
			set
			{
				if ((this._NIP != value))
				{
					this.OnNIPChanging(value);
					this.SendPropertyChanging();
					this._NIP = value;
					this.SendPropertyChanged("NIP");
					this.OnNIPChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.AssociationAttribute(Name="Company_ItemOrder", Storage="_ItemOrders", ThisKey="Id", OtherKey="CompanyId")]
		public EntitySet<ItemOrder> ItemOrders
		{
			get
			{
				return this._ItemOrders;
			}
			set
			{
				this._ItemOrders.Assign(value);
			}
		}
		
		public event PropertyChangingEventHandler PropertyChanging;
		
		public event PropertyChangedEventHandler PropertyChanged;
		
		protected virtual void SendPropertyChanging()
		{
			if ((this.PropertyChanging != null))
			{
				this.PropertyChanging(this, emptyChangingEventArgs);
			}
		}
		
		protected virtual void SendPropertyChanged(String propertyName)
		{
			if ((this.PropertyChanged != null))
			{
				this.PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
			}
		}
		
		private void attach_ItemOrders(ItemOrder entity)
		{
			this.SendPropertyChanging();
			entity.Company = this;
		}
		
		private void detach_ItemOrders(ItemOrder entity)
		{
			this.SendPropertyChanging();
			entity.Company = null;
		}
	}
	
	[global::System.Data.Linq.Mapping.TableAttribute(Name="dbo.InnerOrders")]
	public partial class InnerOrder : INotifyPropertyChanging, INotifyPropertyChanged
	{
		
		private static PropertyChangingEventArgs emptyChangingEventArgs = new PropertyChangingEventArgs(String.Empty);
		
		private int _Id;
		
		private int _ItemId;
		
		private int _CountIn;
		
		private System.DateTime _DateIn;
		
		private int _OrderId;
		
		private EntityRef<Item> _Item;
		
		private EntityRef<ItemOrder> _ItemOrder;
		
    #region Extensibility Method Definitions
    partial void OnLoaded();
    partial void OnValidate(System.Data.Linq.ChangeAction action);
    partial void OnCreated();
    partial void OnIdChanging(int value);
    partial void OnIdChanged();
    partial void OnItemIdChanging(int value);
    partial void OnItemIdChanged();
    partial void OnCountInChanging(int value);
    partial void OnCountInChanged();
    partial void OnDateInChanging(System.DateTime value);
    partial void OnDateInChanged();
    partial void OnOrderIdChanging(int value);
    partial void OnOrderIdChanged();
    #endregion
		
		public InnerOrder()
		{
			this._Item = default(EntityRef<Item>);
			this._ItemOrder = default(EntityRef<ItemOrder>);
			OnCreated();
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_Id", AutoSync=AutoSync.OnInsert, DbType="Int NOT NULL IDENTITY", IsPrimaryKey=true, IsDbGenerated=true)]
		public int Id
		{
			get
			{
				return this._Id;
			}
			set
			{
				if ((this._Id != value))
				{
					this.OnIdChanging(value);
					this.SendPropertyChanging();
					this._Id = value;
					this.SendPropertyChanged("Id");
					this.OnIdChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_ItemId", DbType="Int NOT NULL")]
		public int ItemId
		{
			get
			{
				return this._ItemId;
			}
			set
			{
				if ((this._ItemId != value))
				{
					if (this._Item.HasLoadedOrAssignedValue)
					{
						throw new System.Data.Linq.ForeignKeyReferenceAlreadyHasValueException();
					}
					this.OnItemIdChanging(value);
					this.SendPropertyChanging();
					this._ItemId = value;
					this.SendPropertyChanged("ItemId");
					this.OnItemIdChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_CountIn", DbType="Int NOT NULL")]
		public int CountIn
		{
			get
			{
				return this._CountIn;
			}
			set
			{
				if ((this._CountIn != value))
				{
					this.OnCountInChanging(value);
					this.SendPropertyChanging();
					this._CountIn = value;
					this.SendPropertyChanged("CountIn");
					this.OnCountInChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_DateIn", DbType="Date NOT NULL")]
		public System.DateTime DateIn
		{
			get
			{
				return this._DateIn;
			}
			set
			{
				if ((this._DateIn != value))
				{
					this.OnDateInChanging(value);
					this.SendPropertyChanging();
					this._DateIn = value;
					this.SendPropertyChanged("DateIn");
					this.OnDateInChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_OrderId", DbType="Int NOT NULL")]
		public int OrderId
		{
			get
			{
				return this._OrderId;
			}
			set
			{
				if ((this._OrderId != value))
				{
					if (this._ItemOrder.HasLoadedOrAssignedValue)
					{
						throw new System.Data.Linq.ForeignKeyReferenceAlreadyHasValueException();
					}
					this.OnOrderIdChanging(value);
					this.SendPropertyChanging();
					this._OrderId = value;
					this.SendPropertyChanged("OrderId");
					this.OnOrderIdChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.AssociationAttribute(Name="Item_InnerOrder", Storage="_Item", ThisKey="ItemId", OtherKey="Id", IsForeignKey=true)]
		public Item Item
		{
			get
			{
				return this._Item.Entity;
			}
			set
			{
				Item previousValue = this._Item.Entity;
				if (((previousValue != value) 
							|| (this._Item.HasLoadedOrAssignedValue == false)))
				{
					this.SendPropertyChanging();
					if ((previousValue != null))
					{
						this._Item.Entity = null;
						previousValue.InnerOrders.Remove(this);
					}
					this._Item.Entity = value;
					if ((value != null))
					{
						value.InnerOrders.Add(this);
						this._ItemId = value.Id;
					}
					else
					{
						this._ItemId = default(int);
					}
					this.SendPropertyChanged("Item");
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.AssociationAttribute(Name="ItemOrder_InnerOrder", Storage="_ItemOrder", ThisKey="OrderId", OtherKey="Id", IsForeignKey=true)]
		public ItemOrder ItemOrder
		{
			get
			{
				return this._ItemOrder.Entity;
			}
			set
			{
				ItemOrder previousValue = this._ItemOrder.Entity;
				if (((previousValue != value) 
							|| (this._ItemOrder.HasLoadedOrAssignedValue == false)))
				{
					this.SendPropertyChanging();
					if ((previousValue != null))
					{
						this._ItemOrder.Entity = null;
						previousValue.InnerOrders.Remove(this);
					}
					this._ItemOrder.Entity = value;
					if ((value != null))
					{
						value.InnerOrders.Add(this);
						this._OrderId = value.Id;
					}
					else
					{
						this._OrderId = default(int);
					}
					this.SendPropertyChanged("ItemOrder");
				}
			}
		}
		
		public event PropertyChangingEventHandler PropertyChanging;
		
		public event PropertyChangedEventHandler PropertyChanged;
		
		protected virtual void SendPropertyChanging()
		{
			if ((this.PropertyChanging != null))
			{
				this.PropertyChanging(this, emptyChangingEventArgs);
			}
		}
		
		protected virtual void SendPropertyChanged(String propertyName)
		{
			if ((this.PropertyChanged != null))
			{
				this.PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
			}
		}
	}
	
	[global::System.Data.Linq.Mapping.TableAttribute(Name="dbo.ItemOrders")]
	public partial class ItemOrder : INotifyPropertyChanging, INotifyPropertyChanged
	{
		
		private static PropertyChangingEventArgs emptyChangingEventArgs = new PropertyChangingEventArgs(String.Empty);
		
		private int _Id;
		
		private int _ItemId;
		
		private System.DateTime _DateOut;
		
		private int _CountOut;
		
		private System.Nullable<System.DateTime> _DateIn;
		
		private System.Nullable<int> _CountIn;
		
		private System.Nullable<int> _CountFails;
		
		private System.Nullable<double> _TotalDm;
		
		private bool _Finished;
		
		private bool _PrintList;
		
		private System.Nullable<int> _CompanyId;
		
		private string _Description;
		
		private EntitySet<InnerOrder> _InnerOrders;
		
		private EntityRef<Company> _Company;
		
		private EntityRef<Item> _Item;
		
    #region Extensibility Method Definitions
    partial void OnLoaded();
    partial void OnValidate(System.Data.Linq.ChangeAction action);
    partial void OnCreated();
    partial void OnIdChanging(int value);
    partial void OnIdChanged();
    partial void OnItemIdChanging(int value);
    partial void OnItemIdChanged();
    partial void OnDateOutChanging(System.DateTime value);
    partial void OnDateOutChanged();
    partial void OnCountOutChanging(int value);
    partial void OnCountOutChanged();
    partial void OnDateInChanging(System.Nullable<System.DateTime> value);
    partial void OnDateInChanged();
    partial void OnCountInChanging(System.Nullable<int> value);
    partial void OnCountInChanged();
    partial void OnCountFailsChanging(System.Nullable<int> value);
    partial void OnCountFailsChanged();
    partial void OnTotalDmChanging(System.Nullable<double> value);
    partial void OnTotalDmChanged();
    partial void OnFinishedChanging(bool value);
    partial void OnFinishedChanged();
    partial void OnPrintListChanging(bool value);
    partial void OnPrintListChanged();
    partial void OnCompanyIdChanging(System.Nullable<int> value);
    partial void OnCompanyIdChanged();
    partial void OnDescriptionChanging(string value);
    partial void OnDescriptionChanged();
    #endregion
		
		public ItemOrder()
		{
			this._InnerOrders = new EntitySet<InnerOrder>(new Action<InnerOrder>(this.attach_InnerOrders), new Action<InnerOrder>(this.detach_InnerOrders));
			this._Company = default(EntityRef<Company>);
			this._Item = default(EntityRef<Item>);
			OnCreated();
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_Id", AutoSync=AutoSync.OnInsert, DbType="Int NOT NULL IDENTITY", IsPrimaryKey=true, IsDbGenerated=true)]
		public int Id
		{
			get
			{
				return this._Id;
			}
			set
			{
				if ((this._Id != value))
				{
					this.OnIdChanging(value);
					this.SendPropertyChanging();
					this._Id = value;
					this.SendPropertyChanged("Id");
					this.OnIdChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_ItemId", DbType="Int NOT NULL")]
		public int ItemId
		{
			get
			{
				return this._ItemId;
			}
			set
			{
				if ((this._ItemId != value))
				{
					if (this._Item.HasLoadedOrAssignedValue)
					{
						throw new System.Data.Linq.ForeignKeyReferenceAlreadyHasValueException();
					}
					this.OnItemIdChanging(value);
					this.SendPropertyChanging();
					this._ItemId = value;
					this.SendPropertyChanged("ItemId");
					this.OnItemIdChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_DateOut", DbType="Date NOT NULL")]
		public System.DateTime DateOut
		{
			get
			{
				return this._DateOut;
			}
			set
			{
				if ((this._DateOut != value))
				{
					this.OnDateOutChanging(value);
					this.SendPropertyChanging();
					this._DateOut = value;
					this.SendPropertyChanged("DateOut");
					this.OnDateOutChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_CountOut", DbType="Int NOT NULL")]
		public int CountOut
		{
			get
			{
				return this._CountOut;
			}
			set
			{
				if ((this._CountOut != value))
				{
					this.OnCountOutChanging(value);
					this.SendPropertyChanging();
					this._CountOut = value;
					this.SendPropertyChanged("CountOut");
					this.OnCountOutChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_DateIn", DbType="Date")]
		public System.Nullable<System.DateTime> DateIn
		{
			get
			{
				return this._DateIn;
			}
			set
			{
				if ((this._DateIn != value))
				{
					this.OnDateInChanging(value);
					this.SendPropertyChanging();
					this._DateIn = value;
					this.SendPropertyChanged("DateIn");
					this.OnDateInChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_CountIn", DbType="Int")]
		public System.Nullable<int> CountIn
		{
			get
			{
				return this._CountIn;
			}
			set
			{
				if ((this._CountIn != value))
				{
					this.OnCountInChanging(value);
					this.SendPropertyChanging();
					this._CountIn = value;
					this.SendPropertyChanged("CountIn");
					this.OnCountInChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_CountFails", DbType="Int")]
		public System.Nullable<int> CountFails
		{
			get
			{
				return this._CountFails;
			}
			set
			{
				if ((this._CountFails != value))
				{
					this.OnCountFailsChanging(value);
					this.SendPropertyChanging();
					this._CountFails = value;
					this.SendPropertyChanged("CountFails");
					this.OnCountFailsChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_TotalDm", DbType="Float")]
		public System.Nullable<double> TotalDm
		{
			get
			{
				return this._TotalDm;
			}
			set
			{
				if ((this._TotalDm != value))
				{
					this.OnTotalDmChanging(value);
					this.SendPropertyChanging();
					this._TotalDm = value;
					this.SendPropertyChanged("TotalDm");
					this.OnTotalDmChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_Finished", DbType="Bit NOT NULL")]
		public bool Finished
		{
			get
			{
				return this._Finished;
			}
			set
			{
				if ((this._Finished != value))
				{
					this.OnFinishedChanging(value);
					this.SendPropertyChanging();
					this._Finished = value;
					this.SendPropertyChanged("Finished");
					this.OnFinishedChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_PrintList", DbType="Bit NOT NULL")]
		public bool PrintList
		{
			get
			{
				return this._PrintList;
			}
			set
			{
				if ((this._PrintList != value))
				{
					this.OnPrintListChanging(value);
					this.SendPropertyChanging();
					this._PrintList = value;
					this.SendPropertyChanged("PrintList");
					this.OnPrintListChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_CompanyId", DbType="Int")]
		public System.Nullable<int> CompanyId
		{
			get
			{
				return this._CompanyId;
			}
			set
			{
				if ((this._CompanyId != value))
				{
					if (this._Company.HasLoadedOrAssignedValue)
					{
						throw new System.Data.Linq.ForeignKeyReferenceAlreadyHasValueException();
					}
					this.OnCompanyIdChanging(value);
					this.SendPropertyChanging();
					this._CompanyId = value;
					this.SendPropertyChanged("CompanyId");
					this.OnCompanyIdChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_Description", DbType="NVarChar(MAX)")]
		public string Description
		{
			get
			{
				return this._Description;
			}
			set
			{
				if ((this._Description != value))
				{
					this.OnDescriptionChanging(value);
					this.SendPropertyChanging();
					this._Description = value;
					this.SendPropertyChanged("Description");
					this.OnDescriptionChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.AssociationAttribute(Name="ItemOrder_InnerOrder", Storage="_InnerOrders", ThisKey="Id", OtherKey="OrderId")]
		public EntitySet<InnerOrder> InnerOrders
		{
			get
			{
				return this._InnerOrders;
			}
			set
			{
				this._InnerOrders.Assign(value);
			}
		}
		
		[global::System.Data.Linq.Mapping.AssociationAttribute(Name="Company_ItemOrder", Storage="_Company", ThisKey="CompanyId", OtherKey="Id", IsForeignKey=true)]
		public Company Company
		{
			get
			{
				return this._Company.Entity;
			}
			set
			{
				Company previousValue = this._Company.Entity;
				if (((previousValue != value) 
							|| (this._Company.HasLoadedOrAssignedValue == false)))
				{
					this.SendPropertyChanging();
					if ((previousValue != null))
					{
						this._Company.Entity = null;
						previousValue.ItemOrders.Remove(this);
					}
					this._Company.Entity = value;
					if ((value != null))
					{
						value.ItemOrders.Add(this);
						this._CompanyId = value.Id;
					}
					else
					{
						this._CompanyId = default(Nullable<int>);
					}
					this.SendPropertyChanged("Company");
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.AssociationAttribute(Name="Item_ItemOrder", Storage="_Item", ThisKey="ItemId", OtherKey="Id", IsForeignKey=true)]
		public Item Item
		{
			get
			{
				return this._Item.Entity;
			}
			set
			{
				Item previousValue = this._Item.Entity;
				if (((previousValue != value) 
							|| (this._Item.HasLoadedOrAssignedValue == false)))
				{
					this.SendPropertyChanging();
					if ((previousValue != null))
					{
						this._Item.Entity = null;
						previousValue.ItemOrders.Remove(this);
					}
					this._Item.Entity = value;
					if ((value != null))
					{
						value.ItemOrders.Add(this);
						this._ItemId = value.Id;
					}
					else
					{
						this._ItemId = default(int);
					}
					this.SendPropertyChanged("Item");
				}
			}
		}
		
		public event PropertyChangingEventHandler PropertyChanging;
		
		public event PropertyChangedEventHandler PropertyChanged;
		
		protected virtual void SendPropertyChanging()
		{
			if ((this.PropertyChanging != null))
			{
				this.PropertyChanging(this, emptyChangingEventArgs);
			}
		}
		
		protected virtual void SendPropertyChanged(String propertyName)
		{
			if ((this.PropertyChanged != null))
			{
				this.PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
			}
		}
		
		private void attach_InnerOrders(InnerOrder entity)
		{
			this.SendPropertyChanging();
			entity.ItemOrder = this;
		}
		
		private void detach_InnerOrders(InnerOrder entity)
		{
			this.SendPropertyChanging();
			entity.ItemOrder = null;
		}
	}
}
#pragma warning restore 1591
