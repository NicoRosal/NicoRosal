USE [TECDEV]
GO
/****** Object:  StoredProcedure [dbo].[th_PackingListINCDataRetrieval_Company]    Script Date: 10/11/2021 7:40:25 am ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		<NINO>
-- Create date: <30/06/2021>
-- Description:	<INC Packing List Generate Report>

-- Author:		<GE>
-- Create date: <2021/08/02>
-- Description:	<tmp_NoteContentExpress,tmp_Itemcategory Addition>

-- Author:		<NINO>
-- Create date: <2021/08/02>
-- Description:	<tmp_price_conv,tmp_subtotal,tmp_totalprice>

-- Author:		<NINO>
-- Create date: <2021/08/23>
-- Description:	<tmp_price_conv(2 decimal),tmp_contact, tmp_QtyRemaining (Qty Back Ordered)>

-- Author:		<GE>
-- Create date: <2021/08/24>
-- Description:	<tmp_CustomerLINo>

-- Author:		<NINO>
-- Create date: <2021/08/24>
-- Description:	<ShipCodeDescription, ShipToContact>

-- Author:		<NINO>
-- Create date: <2021/08/26>
-- Description:	<Reworked Join Conditions, custaddr.name>

-- Author:		<NINO>
-- Create date: <2021/09/03>
-- Description:	<Reworked Join Conditions, phone>

-- Author:		<NINO>
-- Create date: <2021/09/08>
-- Description:	<Reworked Join Conditions, NoteExistsFlag>

-- Author:		<GE>
-- Create date: <2021/09/10>
-- Description:	<tmp_account>

-- Author:		<GE>
-- Create date: <2021/09/13>
-- Description:	<ORDER BYはtmp_CustomerLINoを順番にする,QtyRemainingの条件変更>

-- Author:		<Nino>
-- Create date: <2021/09/29>
-- Description:	<Added CTE to Delete Duplicate Entries>

-- Author:		<Nino>
-- Create date: <2021/10/07>
-- Description:	<Added a new item: the connection between tmp_NhNoteFlag and Table. When the internal item is ticked, it is not displayed. 
--				 Modified the expression conditions of Uf_th_CoDropShip, 
--				 the goods must be sent directly to customers and only when they are shipped from the Philippines.>
-- =============================================							--Size	131.375, 27.5
ALTER PROCEDURE [dbo].[th_PackingListINCDataRetrieval]					--Splitter Pos 11.588235294117647
	-- Add the parameters for the stored procedure here
	--Created tmpINCPackingListsIDO | tmpINCPackingList_IDO 2021-07-06 || REMOVED th_PackingListINC Secondary Collections	|| NEW IDO: tmpINCPackingLists || INC Po Number Location: 10.2941176470588, 58.125

		@fac_ship_start				DateType
		,@fac_ship_end				DateType
		,@sls_ship_start			DateType
		,@sls_ship_end				DateType --2021-05-07
		,@cust_po					CustPoType
		,@cust_num					CustNumType	
		,@po_num					PoNumType
		--object.tmp_co_num
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.

	SET NOCOUNT ON;

    -- Insert statements for procedure here
	--SP DATA RETRIEVAL FOR INC PACKING LIST
			--FormRefresh
		--CollectionRefresh
-- =============================================
		SET ANSI_WARNINGS OFF  -- SET ISO Standard off

		SET @fac_ship_start = ISNULL(@fac_ship_start,'')
		SET @fac_ship_end = ISNULL(@fac_ship_end,'')
		SET @sls_ship_start = ISNULL(@sls_ship_start,'')
		SET @sls_ship_end = ISNULL (@sls_ship_end, '') --2021-07-05
		SET @cust_po = ISNULL(@cust_po,'')
		SET @cust_num = ISNULL(@cust_num,'')
		SET @po_num = ISNULL(@po_num,'')

	TRUNCATE TABLE tmpINCPackingList
	INSERT INTO tmpINCPackingList
	(
		tmp_co_num
		,tmp_qty_partial
		,tmp_qty_ordered
		,tmp_qty_shipped
		,tmp_cust_item
		,tmp_QtyRemaining
		,tmp_cust_po
		,tmp_item
		,tmp_ItemDescription
		,tmp_CountryOfOrigin
		,tmp_fac_shipdate_actual
		,tmp_ship_date
		,tmp_CustName
		,tmp_cust_num
		,tmp_carrier_bill_to_transportation
		,tmp_ShipCodeDescription
		,tmp_ShippingTerms
		,tmp_weight
		,tmp_CustomerLINo
		,tmp_fulladdress	--ADDED 12/07/2021
		,tmp_city
		,tmp_state
		,tmp_county
		,tmp_zip
		,tmp_country
		,tmp_CustAddr1
		,tmp_CustAddr2
		,tmp_CustAddr3
		,tmp_ShipToCompanyName_EN	
		,tmp_ShipToContact
		,tmp_phoneno1
		,tmp_NoteExistsFlag
		,tmp_do_num
		,tmp_NoteDesc
		,tmp_NoteContent
		,tmp_carrier_account
		,tmp_Uf_DutyTaxChargeTo
--NEW
		,tmp_fulladdress2
		,tmp_city_2
		,tmp_state_2
		,tmp_county_2
		,tmp_zip_2
		,tmp_country_2
		,tmp_custaddr01_2
		,tmp_custaddr02_2
		,tmp_custaddr03_2
		,tmp_CompanyNameEN_2
		,tmp_phoneno1_2
		,tmp_boxsize
		,tmp_Uf_th_CoDropShip
		,tmp_po_num
		,tmp_datepacked
		-- 2021/08/02 GE add　Start --
		,tmp_NoteContentExpress
		,tmp_Itemcategory
		-- 2021/08/02 GE add　End --
		-- 2021/08/05 Nino add　Start --
		,tmp_price_conv
		,tmp_subtotal
		,tmp_totalprice
		-- 2021/08/05 Nino add　End --
		-- 2021/08/23 Nino add　Start --
		,tmp_contact
		-- 2021/08/23 Nino add　End --
		-- 2021/09/10 GE add　Start --
		,tmp_account
		-- 2021/09/10 GE add　End --
		-- 2021/09/21 Nino add　Start --
		,tmp_partial_num
		,tmp_do_hdr_date
		,tmp_ref_num
		-- 2021/09/21 Nino add　End --
		-- 2021/10/07 GE Add　Start --
		,tmp_NhNoteFlag
		-- 2021/10/07 GE Add　End --
		)

	SELECT 
			co_num
			,qty_partial
			,qty_ordered
			,qty_shipped
			,cust_item
			,QtyRemaining
			,cust_po
			,item
			,ItemDescription
			,CountryOfOrigin
			,fac_shipdate_plan
			,sc_shipdate_plan
			,CustomerName
			,cust_num
			,carrier_bill_to_transportation
			,ShipCodeDescription
			,ShippingTerms
			-- 2021/09/17 Nino add　Start --
			,weight
		--	,NetWeight
			-- 2021/09/17 Nino add　End --
			,CustomerLI#
			--,fulladdress
			,CASE WHEN COALESCE(CustAddr2, CustAddr3,fulladdress) = CustAddr2
				THEN (CASE WHEN COALESCE(CustAddr3,fulladdress) = CustAddr3
					THEN fulladdress
					WHEN  COALESCE(CustAddr3,fulladdress) = fulladdress
						THEN NULL
							END)
					WHEN COALESCE(CustAddr2, CustAddr3,fulladdress) = CustAddr3
						THEN NULL
					WHEN COALESCE(CustAddr2, CustAddr3,fulladdress) = fulladdress
						THEN NULL
							END
			,city
			,state
			,county
			,zip
			,country
			,ISNULL(CustAddr1, 'No Address[1]')	

			,COALESCE(CustAddr2,CustAddr3,fulladdress)
			,CASE WHEN COALESCE(CustAddr2, CustAddr3,fulladdress) = CustAddr2
					THEN (CASE WHEN COALESCE(CustAddr3,fulladdress) = CustAddr3
							THEN CustAddr3
							WHEN  COALESCE(CustAddr3,fulladdress) = fulladdress
								THEN fulladdress
									END)
					WHEN COALESCE(CustAddr2, CustAddr3,fulladdress) = CustAddr3
						THEN (CASE WHEN COALESCE(CustAddr3,fulladdress) = CustAddr3
							THEN fulladdress
									END)
					WHEN COALESCE(CustAddr2, CustAddr3,fulladdress) = fulladdress
						THEN NULL
									END



			,CompanyNameEN	
			,ShipToContact
			,phone##1
			,NoteExistsFlag
			,do_num
			,ISNULL(NoteDesc, NULL) AS NoteDesc
			,ISNULL(NoteContent, NULL	) AS NoteContent
	--		,CarrierAccount
			,CASE WHEN carrieraccount LIKE '/%'
				THEN RIGHT(carrieraccount, LEN(carrieraccount)-1)
					ELSE carrieraccount
				END AS carrieraccount	
			,DutyTaxChargeTo
--NEW
			--,fulladdress2
			,CASE WHEN COALESCE(custaddr02_2,custaddr03_2,fulladdress2,phone##1_2) = custaddr02_2
				THEN (CASE WHEN COALESCE(custaddr03_2,fulladdress2,phone##1_2) = custaddr03_2
							THEN COALESCE(fulladdress2,phone##1_2)
						WHEN COALESCE(custaddr03_2,fulladdress2,phone##1_2) = fulladdress2
							THEN (CASE WHEN COALESCE(fulladdress2,phone##1_2) = fulladdress2
										THEN phone##1_2
									WHEN COALESCE(fulladdress2,phone##1_2) = phone##1_2
										THEN NULL
									END)
						WHEN COALESCE(custaddr03_2,fulladdress2,phone##1_2) = phone##1_2
							THEN NULL
								END)
				WHEN COALESCE(custaddr02_2,custaddr03_2,fulladdress2,phone##1_2) = custaddr03_2
					THEN (CASE WHEN COALESCE(fulladdress2,phone##1_2) = fulladdress2
						THEN phone##1_2
							WHEN COALESCE(fulladdress2,phone##1_2) = phone##1_2
								THEN NULL
									END)
				WHEN COALESCE(custaddr02_2,custaddr03_2,fulladdress2,phone##1_2) = fulladdress2
					THEN NULL
								END AS fulladdress2
			,city_2			--REVISED 19/07/2021
			,state_2
			,county_2
			,ZIP_2
			,country2
			,ISNULL(custaddr01_2, 'No Address[1]')
			,COALESCE(custaddr02_2,custaddr03_2,fulladdress2,phone##1_2)
			,CASE WHEN COALESCE(custaddr02_2,custaddr03_2,fulladdress2,phone##1_2) = custaddr02_2
				THEN (CASE WHEN COALESCE(custaddr03_2,fulladdress2,phone##1_2) = custaddr03_2
					THEN custaddr03_2
						WHEN COALESCE(custaddr03_2,fulladdress2,phone##1_2) = fulladdress2
							THEN fulladdress2
								WHEN COALESCE(custaddr03_2,fulladdress2,phone##1_2) = phone##1_2
									THEN phone##1_2
										END)
				WHEN COALESCE(custaddr02_2,custaddr03_2,fulladdress2,phone##1_2) = custaddr03_2
					THEN (CASE WHEN COALESCE(custaddr03_2,fulladdress2,phone##1_2) = custaddr03_2
						THEN COALESCE(fulladdress2,phone##1_2)
							END)
				WHEN COALESCE(custaddr02_2,custaddr03_2,fulladdress2,phone##1_2) = fulladdress2
					THEN (CASE WHEN COALESCE(fulladdress2,phone##1_2) = fulladdress2
						THEN phone##1_2
							END)
				WHEN COALESCE(custaddr02_2,custaddr03_2,fulladdress2,phone##1_2) = phone##1_2
					THEN NULL
					END
			,CompanyNameEN_2
			,CASE WHEN COALESCE(custaddr02_2,custaddr03_2,fulladdress2,phone##1_2) = custaddr02_2 
				THEN (CASE WHEN COALESCE(custaddr03_2,fulladdress2,phone##1_2) = custaddr03_2
					THEN (CASE WHEN COALESCE(fulladdress2,phone##1_2) = fulladdress2
						THEN phone##1_2
							WHEN COALESCE(fulladdress2,phone##1_2) = phone##1_2
								THEN NULL
								END)
						WHEN COALESCE(custaddr03_2,fulladdress2,phone##1_2) = fulladdress2
							THEN NULL
						WHEN COALESCE(custaddr03_2,fulladdress2,phone##1_2) = phone##1_2
							THEN NULL
								END)
				WHEN COALESCE(custaddr02_2,custaddr03_2,fulladdress2,phone##1_2) = custaddr03_2 
					THEN NULL
				WHEN COALESCE(custaddr02_2,custaddr03_2,fulladdress2,phone##1_2) = fulladdress2 
					THEN NULL
				WHEN COALESCE(custaddr02_2,custaddr03_2,fulladdress2,phone##1_2) = phone##1_2 
					THEN NULL
					END
				,description
				-- 2021/10/07 GE add　Start --
				--,CASE WHEN Uf_th_CoDropShip = 1 
				--	THEN 'Drop shipped from our facility in Cebu Philippines'	--YES		
				--WHEN Uf_th_CoDropShip = 2
				--		THEN NULL	--NO
				--	ELSE ISNULL(Uf_th_CoDropShip, NULL)
				--	END AS Uf_th_CoDropShip 
				,Uf_th_CoDropShip
				-- 2021/10/07 GE add　End --
				,po_num
				,CASE WHEN CAST(fac_shipdate_plan AS NVARCHAR(30)) IS NULL	
					THEN 'Date Packed: '
						ELSE 'Date Packed: ' + CAST(fac_shipdate_plan AS NVARCHAR(30))
						END

			-- 2021/08/02 GE add　Start --
				,CASE WHEN NoteDesc IS NOT NULL THEN 1 ELSE 0 END
				,itemcategory
			-- 2021/08/02 GE add　End --
			-- 2021/08/05 Nino add　Start --
				-- 2021/08/23 Nino add　Start --
				,FORMAT(price_conv, '#,0.00') AS Price	--Changed to Two Decimal Places 2021/08/23
				-- 2021/08/23 Nino add　End --
				,FORMAT(price_conv * qty_ordered, '#,0.00') AS SubTotal
				-- 2021/09/16 Nino add　Start --
		--		,FORMAT(CAST(SUM(price_conv * qty_ordered)  OVER(PARTITION BY getdate()) AS decimal(23,2)),'#,0.00') AS Total
				,FORMAT(CAST(SUM(price_conv * qty_ordered)  OVER(PARTITION BY Uf_th_CoDropShip) AS decimal(23,2)),'#,0.00') AS Total
				-- 2021/09/16 Nino add　End --
			-- 2021/08/05 Nino add　End --
			-- 2021/08/23 Nino add　Start --
			,contact
			-- 2021/08/23 Nino add　End --
			-- 2021/09/10 GE add　Start --
			,CASE WHEN Account LIKE '/%'
				THEN RIGHT(Account, LEN(Account)-1)
					ELSE Account
				END AS Account
			-- 2021/09/10 GE add　End --
			-- 2021/09/21 Nino add　Start --
			,partial_num
			,do_hdr_date
			,ref_num
			-- 2021/09/21 Nino add　End --
			-- 2021/10/07 GE Add　Start --
			,NoteFlag
			-- 2021/10/07 GE Add　End --
	FROM --INCPackingList_view
			(
			SELECT DISTINCT

				co.co_num			
				,CAST(mdd.qty_partial AS INT) AS qty_partial
				,CAST(coitem.qty_ordered AS INT) AS qty_ordered
				,CAST(coitem.qty_shipped AS INT) AS qty_shipped
				,coitem.cust_item
				-- GE Add 2021/09/13 Start --
				-- 2021/08/23 Nino add　Start --
				--,ISNULL(coitem.qty_ordered ,'0') - ISNULL(coitem.qty_shipped, '0') AS QtyRemaining
				-- 2021/08/23 Nino add　End --
				,CASE WHEN SUM(mdd.partial_num) OVER(PARTITION BY co.co_num, co.cust_po, coitem.price_conv, coitem.item, poi.po_num ORDER BY co.co_num, coitem.item, mdd.partial_num) >= 3 THEN
					CAST(ISNULL(coitem.qty_ordered , 0) - (SUM(ISNULL(mdd.qty_partial, 0))  OVER(PARTITION BY 
																			co.co_num
																			,co.cust_po
																			,coitem.price_conv
																			,ISNULL(coitem.item, '')
																			,ISNULL(coitem.description, '')
																			,poi.po_num
																			
																			,CAST(ISNULL(sn.NoteDesc, '') AS NVARCHAR(100))
																			,CAST(ISNULL(sn.NoteContent, '') AS NVARCHAR(100))
																			,fdol.weight
																			,fdol.description

																			ORDER BY 
																			mdd.partial_num)) AS INT)  

						ELSE 
							CAST(ISNULL(coitem.qty_ordered, 0) - ISNULL(mdd.qty_partial, 0) AS INT)
						  END 
						  AS QtyRemaining
				-- GE Add 2021/09/13 End --
				,ISNULL(co.cust_po,'') AS cust_po	
				,ISNULL(coitem.item, '') AS item
				,ISNULL(coitem.description, '') AS ItemDescription
				,ISNULL(i.country,'') AS CountryOfOrigin
				,CAST(mdd.fac_shipdate_plan AS date) AS fac_shipdate_plan
				,CAST(mdd.sc_shipdate_plan AS date) AS sc_shipdate_plan	
				-- 2021/08/26 Nino add　Start --
				,ISNULL(custo.Uf_th_cusname1, NULL) AS CustomerName
				-- 2021/08/26 Nino add　End --
				,custo.cust_num
				,CASE WHEN custaddr0.carrier_bill_to_transportation = 'S'
					THEN 'Shipper'
					WHEN custaddr0.carrier_bill_to_transportation = 'R'
						THEN 'Receiver'
					WHEN custaddr0.carrier_bill_to_transportation = 'T' OR custaddr0.carrier_bill_to_transportation LIKE 'T%'
						THEN 'Third-Party'
					ELSE ISNULL(custaddr0.carrier_bill_to_transportation, '')
					END AS carrier_bill_to_transportation
				,ISNULL(ship.description, '') AS ShipCodeDescription
				,ISNULL(custo.Uf_th_incoterms,'') AS ShippingTerms
				-- 2021/09/21 Nino add　Start -- 
				,fdol.weight --added 09/07/2021	--CAUSING DUPLICATES
			--	,seq.NetWeight
				-- 2021/09/21 Nino add　End -- 
				--,ISNULL(coitem.co_line,'') AS CustomerLI#
				,ISNULL(coitem.Uf_InvoiceNoteInc,'') AS CustomerLI#
				-- 2021/08/24 GE add　End --
				
				,CAST(
					CONCAT(
						CASE custaddr.city
							WHEN NULL THEN ''
							ELSE custaddr.city + ', '
							END
						,CASE custaddr.state
							WHEN NULL THEN ''
							ELSE custaddr.state + ', '
							END
						,CASE custaddr.county
							WHEN NULL THEN ''
							ELSE custaddr.county + ', '
							END
						,CASE custaddr.zip
							WHEN NULL THEN ''
							ELSE custaddr.zip + ', '		
							END
						,CASE custaddr.country
							WHEN NULL THEN ''
							WHEN (
								CASE WHEN custaddr.zip LIKE '%'+',' AND custaddr.country IS NULL THEN 
									 LEFT(custaddr.zip, LEN(custaddr.zip) - 1)
									ELSE custaddr.zip
									END
							)
								THEN ''
							ELSE custaddr.country
							END
							) AS NVARCHAR(120))
								AS fulladdress
				,ISNULL(custaddr.city,NULL) AS city
				,ISNULL(custaddr.state,NULL) AS state
				,ISNULL(custaddr.county,NULL) AS county
				,ISNULL(custaddr.zip,NULL) AS ZIP
				,ISNULL(custaddr.country,NULL) AS country
				,ISNULL(custaddr.addr##1,NULL) AS CustAddr1
				,ISNULL(custaddr.addr##2,NULL) AS CustAddr2
				,ISNULL(custo.Uf_th_cusadd3,NULL) AS CustAddr3
				-- 2021/08/26 Nino add　Start --
				,ISNULL(custaddr.name, NULL) AS CompanyNameEN
				-- 2021/08/26 Nino add　End --
				--,ISNULL(custo.Uf_th_cusname1,NULL) AS CompanyNameEN
				-- 2021/08/24 Nino add　Start --
				,ISNULL(co.Uf_th_ShiptoContact,'No Contact Person Available') AS ShipToContact
				-- 2021/08/24 Nino add　End --
				-- 2021/09/03 Nino add　Start --
				--,ISNULL(custo.phone##1,NULL) AS phone##1
				,ISNULL(custo.phone##2,NULL) AS phone##1
				-- 2021/09/03 Nino add　End --
				-- 2021/09/07 Nino add　Start --
			--	,coitem.NoteExistsFlag
				,CASE WHEN co.Uf_th_CoDropShip = 1
					THEN NULL	--YES
					WHEN co.Uf_th_CoDropShip = 2
						THEN '1'	--NO
					ELSE ISNULL(co.Uf_th_CoDropShip, NULL)
					END AS NoteExistsFlag 
				---- 2021/09/07 Nino add　End --
				,ISNULL(hdr.do_num, '') AS do_num
				,CONCAT(custaddr0.carrier_account, '/'+ ' ',
					CASE WHEN custaddr0.carrier_bill_to_transportation = 'S'
						THEN 'Shipper'
							WHEN custaddr0.carrier_bill_to_transportation = 'R'
								THEN 'Receiver'
							WHEN custaddr0.carrier_bill_to_transportation LIKE 'T%'
								THEN 'Third-Party'
							ELSE custaddr.carrier_bill_to_transportation
							END 
				) AS carrieraccount	--/S
		
				,ISNULL(custo.Uf_DutyTaxChargeTo,'') AS DutyTaxChargeTo
--From Misaki	
				,CONVERT(NVARCHAR(255),sn.NoteDesc)	AS NoteDesc	--ntext DataType conv to NVARCHAR for DISTINCT to work
				,CONVERT(NVARCHAR(255),sn.NoteContent) AS NoteContent	--ntext DataType conv to NVARCHAR for DISTINCT to work
--NEW
				,ISNULL(custaddr0.city,NULL) AS city_2
				,ISNULL(custaddr0.state,NULL) AS state_2
				,ISNULL(custaddr0.county,NULL) AS county_2
				,ISNULL(custaddr0.zip,NULL) AS ZIP_2
				,ISNULL(custaddr0.country,NULL) AS country2
				,custaddr0.addr##1 AS custaddr01_2
				,custaddr0.addr##2 AS custaddr02_2
				,custo0.Uf_th_cusadd3 AS custaddr03_2
				,ISNULL(custo0.Uf_th_cusname1,'No Customer Name') AS CompanyNameEN_2
				-- 2021/09/03 Nino add　Start --
				,ISNULL(custo0.phone##2,NULL) AS phone##1_2
				-- 2021/09/03 Nino add　End --
				,CAST(
					CONCAT(
						CASE custaddr0.city
							WHEN NULL THEN ''
							ELSE custaddr0.city + ', '
							END
						,CASE custaddr0.state
							WHEN NULL THEN ''
							ELSE custaddr0.state + ', '
							END
						,CASE custaddr0.county
							WHEN NULL THEN ''
							ELSE custaddr0.county + ', '
							END
						,CASE custaddr0.zip
							WHEN NULL THEN ''
							ELSE custaddr0.zip + ', '		
							END
						,CASE custaddr0.country
							WHEN NULL THEN ''
							WHEN (
								CASE WHEN custaddr0.zip LIKE '%'+',' AND custaddr0.country IS NULL THEN 
									 LEFT(custaddr0.zip, LEN(custaddr0.zip) - 1)
									ELSE custaddr0.zip
									END
							)
								THEN ''
							ELSE custaddr0.country
							END
							) AS NVARCHAR(120))
								AS fulladdress2
				-- 2021/10/07 GE add　Start --		
				--,co.Uf_th_CoDropShip 
				,CASE WHEN co.Uf_th_CoDropShip = '1' AND custo.Uf_th_incoterms = 'EXW Philippines' OR custo.Uf_th_incoterms = 'FCA Philippines'
					THEN 'Drop shipped from our facility in Cebu Philippines'	--YES		
				WHEN co.Uf_th_CoDropShip = '2'
						THEN NULL	--NO
					--ELSE ISNULL(co.Uf_th_CoDropShip, NULL)
					END AS Uf_th_CoDropShip 
				-- 2021/10/07 GE add　End --
				,fdol.description		--CAUSING DUPLICATES
				,poi.po_num
				-- 2021/08/02 GE add　Start --
				,CASE WHEN item.charfld3 = '2100' THEN '(Chip Capacitor)' 
				WHEN item.charfld3 = '2200' THEN '(Substrate)'
				WHEN item.charfld3 = '2300' THEN '(Chip Resistor)'
				WHEN item.charfld3 = '4310' THEN '(Nozzle)'
				WHEN item.charfld3 = '4550' THEN '(Needle)'
				WHEN item.charfld3 = '4560' THEN '(Diamond Needle)'
				WHEN item.charfld3 = '4700' THEN '(Other Tools)'
				WHEN item.charfld3 = '4730' THEN '(Pinhole CAP Tool)'
				WHEN item.charfld3 = '8100' THEN '(Scriber)' END AS itemcategory
				-- 2021/08/02 GE add　End --
				-- 2021/08/05 Nino add　Start --
				,coitem.price_conv
				-- 2021/08/05 Nino add　End --
				-- 2021/08/23 Nino add　Start --
				,co.contact
				-- 2021/08/23 Nino add　End --
				-- 2021/09/10 GE add　End --
				,custaddr0.carrier_account AS Account
				-- 2021/09/10 GE add　End --
				-- 2021/09/21 Nino add　Start --
				,mdd.partial_num
				,hdr.do_hdr_date
				,fdos.ref_num	--seq.ref_num
				-- 2021/09/21 Nino add　End --
				-- 2021/10/07 GE Add　Start --
				,CASE WHEN nh.NoteFlag = '1' THEN 1
				 ELSE 0 END AS NoteFlag
				 -- 2021/10/07 GE Add　End --


			FROM 	TDH_MngDeliveryDate_mst mdd --co_mst co
				INNER JOIN coitem_mst coitem
					ON coitem.site_ref = mdd.site_ref AND coitem.co_num = mdd.num AND coitem.co_line = mdd.line AND coitem.co_release = mdd.release
				INNER JOIN  co_mst co--TDH_MngDeliveryDate_mst mdd
					ON co.site_ref = coitem.site_ref AND co.co_num = coitem.co_num AND co.site_ref = 'INC' --AND co.stat = coitem.stat 
				LEFT JOIN custaddr_mst AS custaddr
					ON custaddr.site_ref = co.site_ref AND custaddr.cust_num = co.cust_num-- AND custaddr.cust_seq = co.cust_seq
					-- 2021/08/26 Nino add　Start --
						AND custaddr.cust_seq = 0
					-- 2021/08/26 Nino add　End --
				LEFT JOIN state_mst AS st
					ON st.site_ref = co.site_ref AND st.state = custaddr.state
				LEFT JOIN custaddr_mst AS custaddr0
					ON custaddr0.site_ref = co.site_ref AND custaddr0.cust_num = co.cust_num AND custaddr0.cust_seq = co.cust_seq
				LEFT JOIN item_mst i
					ON i.site_ref = coitem.site_ref AND i.item = coitem.item --AND i.p_m_t_code = 'P'
				LEFT JOIN poitem_mst poi
					ON poi.site_ref = coitem.site_ref AND poi.po_num = coitem.ref_num AND poi.po_line = coitem.ref_line_suf AND poi.po_release = coitem.ref_release --AND coitem.ref_type = 'P'
				LEFT JOIN po_mst po
					ON po.site_ref = poi.site_ref AND po.po_num = poi.po_num
				LEFT JOIN vendor_mst ven
					ON ven.site_ref = po.site_ref AND ven.vend_num = po.vend_num
				LEFT JOIN coitem_mst coif
					ON coif.site_ref = ven.source_site AND coif.co_num = po.source_site_co_num AND coif.co_line = poi.po_line --AND coif.co_release = 0
				LEFT JOIN do_seq_mst fdos
					ON fdos.site_ref = coif.site_ref AND fdos.ref_num = coif.co_num AND fdos.ref_line = coif.co_line
				LEFT JOIN do_line_mst fdol
					ON fdos.do_num = fdol.do_num AND fdos.do_line = fdol.do_line
				LEFT JOIN customer_mst custo
					ON custo.site_ref = co.site_ref AND custo.cust_num = co.cust_num 
					AND custo.cust_seq = co.cust_seq--AND custo.Uf_th_incoterms IS NOT NULL--AND custo.cust_seq = 0
				LEFT JOIN customer_mst custo0
					ON custo0.site_ref = co.site_ref AND custo0.cust_num = co.cust_num 
					AND custo0.cust_seq = co.cust_seq-- AND custo0.Uf_th_incoterms IS NOT NULL
				LEFT JOIN state_mst as st0
					ON st0.site_ref = co.site_ref AND st0.state = custaddr0.state
				LEFT JOIN shipcode_mst ship
					ON ship.site_ref = custo0.site_ref AND ship.ship_code = custo0.ship_code
				LEFT JOIN carrier_mst car
					ON car.site_ref = ship.site_ref AND car.carrier_code = ship.carrier_code
				LEFT JOIN ObjectNotes obn
					ON obn.RefRowPointer = co.RowPointer
				-- 2021/10/07 GE Add　Start --
				LEFT JOIN NoteHeaders nh
					ON nh.NoteHeaderToken = obn.NoteHeaderToken	
				-- 2021/10/07 GE Add　End --	
				LEFT JOIN SpecificNotes sn
					ON sn.SpecificNoteToken = obn.SpecificNoteToken		
				LEFT JOIN item_mst fi
					ON fi.site_ref = coif.site_ref AND fi.item = coitem.item
				LEFT JOIN item
					ON item.item = coitem.item
				-- 2021/09/21 Nino add　Start --	
				LEFT JOIN do_hdr_mst hdr
					ON hdr.site_ref = coitem.site_ref AND hdr.do_num = fdos.do_num
				-- 2021/09/21 Nino add　End --
		)
			AS TableTest1

			WHERE			
				((@fac_ship_start = '') OR (fac_shipdate_plan BETWEEN @fac_ship_start AND @fac_ship_end) )
				AND ((@fac_ship_end = '') OR (fac_shipdate_plan BETWEEN @fac_ship_start AND @fac_ship_end) )
				AND ((@sls_ship_start = '') OR (sc_shipdate_plan BETWEEN @sls_ship_start AND @sls_ship_end) )
				AND ((@sls_ship_end = '') OR (sc_shipdate_plan BETWEEN @sls_ship_start AND @sls_ship_end) )
				AND ((@cust_po = '') OR (cust_po = @cust_po) )
				AND ((@cust_num = '') OR (cust_num = @cust_num) )
				AND ((@po_num = '') OR (po_num = @po_num) );

			--DELETE DUPLICATE ROWS
				WITH cte_del AS (
					SELECT
						tmp_co_num
						,tmp_item
						,tmp_partial_num
						,tmp_fac_shipdate_actual
						,tmp_ship_date
						,ROW_NUMBER() OVER(
										PARTITION BY
											tmp_co_num
											,tmp_item
											,tmp_partial_num
											,tmp_fac_shipdate_actual
											,tmp_ship_date
											,tmp_CustomerLINo
										ORDER BY
											tmp_co_num
											,tmp_item
											,tmp_partial_num
											,tmp_fac_shipdate_actual
											,tmp_ship_date
												) row_num 
													FROM
														tmpINCPackingList
														)
														DELETE FROM cte_del 
														WHERE row_num > 1;
																																									
	SET ANSI_WARNINGS ON

					SELECT
						tmp_co_num
						,tmp_qty_partial
						,tmp_qty_ordered
						,tmp_qty_shipped
						,tmp_cust_item
						,tmp_QtyRemaining
						,tmp_cust_po
						,tmp_item
						,tmp_ItemDescription
						,tmp_CountryOfOrigin
						,tmp_fac_shipdate_actual
						,tmp_ship_date
						,tmp_CustName
						,tmp_cust_num
						,tmp_carrier_bill_to_transportation
						,tmp_ShipCodeDescription
						,tmp_ShippingTerms
						,tmp_weight
						,tmp_CustomerLINo
						,CASE WHEN tmp_fulladdress LIKE '%'+','+ ' '
								THEN LEFT(tmp_fulladdress, LEN(tmp_fulladdress)-1)
								ELSE tmp_fulladdress
							END AS tmp_fulladdress --ADDED 12/07/2021 REVISED 13/07/2021
						,tmp_city
						,tmp_state
						,tmp_county
						,tmp_zip
						,tmp_country
						,tmp_CustAddr1
						,tmp_CustAddr2
						,tmp_CustAddr3
						,tmp_ShipToCompanyName_EN	
						,tmp_ShipToContact
						,tmp_phoneno1
						,tmp_NoteExistsFlag
						,tmp_do_num
						,tmp_NoteDesc
						,tmp_NoteContent
						,CASE WHEN tmp_carrier_account LIKE '/%' AND tmp_carrier_bill_to_transportation IS NOT NULL
							THEN RIGHT(tmp_carrier_account, LEN(tmp_carrier_account)-1)
								ELSE tmp_carrier_account	
							END AS tmp_carrier_account	
						,tmp_Uf_DutyTaxChargeTo
						,CASE WHEN tmp_fulladdress2 LIKE '%' + ',' + ' '
							THEN LEFT(tmp_fulladdress2, LEN(tmp_fulladdress2)-1)
								WHEN tmp_fulladdress2 IS NULL OR tmp_fulladdress2 = ''
									THEN tmp_phoneno1_2
							ELSE tmp_fulladdress2
							END AS tmp_fulladdress2
						,tmp_city_2
						,tmp_state_2
						,tmp_county_2
						,tmp_zip_2
						,tmp_country_2
						,tmp_custaddr01_2
						,tmp_custaddr02_2
						,tmp_custaddr03_2
						,tmp_CompanyNameEN_2
						,tmp_phoneno1_2
						,tmp_boxsize
						,tmp_Uf_th_CoDropShip
						,tmp_po_num
						,tmp_datepacked
						-- 2021/08/02 GE add　Start --
						,tmp_NoteContentExpress
						,tmp_Itemcategory
						-- 2021/08/02 GE add　End --
						-- 2021/08/05 Nino add　Start --
						,tmp_price_conv
						,tmp_subtotal
						,tmp_totalprice
						-- 2021/08/05 Nino add　End --
						-- 2021/08/23 Nino add　Start --
						,tmp_contact
						-- 2021/08/23 Nino add　End --
						-- 2021/09/10 GE add　Start --
						,tmp_account
						-- 2021/09/10 GE add　End --
						-- 2021/09/21 Nino add　Start --
						,tmp_partial_num
						,tmp_do_hdr_date
						,tmp_ref_num
						-- 2021/09/21 Nino add　End --
						-- 2021/10/07 GE Add　Start --
						,tmp_NhNoteFlag
						-- 2021/10/07 GE Add　End --
					FROM tmpINCPackingList --'INC Packing'
						-- GE Add 2021/09/13 Start --
						--ORDER BY tmp_CustName DESC
						ORDER BY tmp_CustName DESC,tmp_CustomerLINo
								,tmp_co_num, tmp_partial_num
						-- GE Add 2021/09/13 End --



						/*
						--From Norwin To Get More Rows
    DECLARE @Severity    INT = 0
    DECLARE @Infobar    InfobarType;
	
    SELECT
       N' SELECT * ' AS SelectionClause
     , N' FROM tmpINCPackingList '      AS FromClause
     , N' WHERE 1=1 '                        AS WhereClause
     , N' ORDER BY tmp_fac_shipdate_actual '            AS AdditionalClause
     , N''        AS KeyColumns
     , N''                                AS FilterString
    INTO #DynamicParameters
    
    EXEC @Severity = dbo.ExecuteDynamicSQLSp
         @NeedGetMoreRows  = 1
       , @Infobar          = @Infobar          OUTPUT

    RETURN @Severity
*/

END	

