|Object| What is new| Description|Feedback|
|:----|:----|:----|:----|
|[application](reference/excel/application.md)|_Property_ > calculationEngineVersion|Returns a number about the version of Excel Calculation Engine that the workbook was last fully recalculated by. Read-only.|beta|
|[application](reference/excel/application.md)|_Relationship_ > calculationState|Returns a CalculationState that indicates the calculation state of the application. Read-only.|beta|
|[application](reference/excel/application.md)|_Relationship_ > iterativeCalculation|Returns the Iterative Calculation settings. Read-only.|beta|
|[application](reference/excel/application.md)|_Method_ > [createWorkbook(base64File: string)]((reference/excel/application.md#createworkbookbase64file-string)|Creates a new hidden workbook by using an optional base64 encoded .xlsx file.|1.8|
|[application](reference/excel/application.md)|_Method_ > [suspendScreenUpdatingUntilNextSync()]((reference/excel/application.md#suspendscreenupdatinguntilnextsync)|Suspends sceen updating until the next "context.sync()" is called.|beta|
|[autoFilter](reference/excel/autofilter.md)|_Property_ > enabled|Indicates if the AutoFilter is enabled or not. Read-Only. Read-only.|beta|
|[autoFilter](reference/excel/autofilter.md)|_Property_ > isDataFiltered|Indicates if the AutoFilter has filter criteria. Read-Only. Read-only.|beta|
|[autoFilter](reference/excel/autofilter.md)|_Relationship_ > criteria|Array that holds all filter criterias in an autofiltered range. Read-Only. Read-only.|beta|
|[autoFilter](reference/excel/autofilter.md)|_Method_ > [apply(range: Range or string, columnIndex: number, criteria: FilterCriteria)]((reference/excel/autofilter.md#applyrange-range-or-string-columnindex-number-criteria-filtercriteria)|Applies AutoFilter on a range and filters the column if column index and filter criteria are specified.|beta|
|[autoFilter](reference/excel/autofilter.md)|_Method_ > [clearCriteria()]((reference/excel/autofilter.md#clearcriteria)|Clears the criteria if AutoFilter has filters|beta|
|[autoFilter](reference/excel/autofilter.md)|_Method_ > [getRange()]((reference/excel/autofilter.md#getrange)|Returns the Range object that represents the range to which the AutoFilter applies.|beta|
|[autoFilter](reference/excel/autofilter.md)|_Method_ > [getRangeOrNullObject()]((reference/excel/autofilter.md#getrangeornullobject)|If there is Range object associated with the AutoFilter, this method returns it.|beta|
|[autoFilter](reference/excel/autofilter.md)|_Method_ > [reapply()]((reference/excel/autofilter.md#reapply)|Applies the specified Autofilter object currently on the range.|beta|
|[autoFilter](reference/excel/autofilter.md)|_Method_ > [remove()]((reference/excel/autofilter.md#remove)|Removes the AutoFilter for the range.|beta|
|[basicDataValidation](reference/excel/basicdatavalidation.md)|_Property_ > formula1|Gets or sets the Formula1, i.e. minimum value or value depending of the operator.|1.8|
|[basicDataValidation](reference/excel/basicdatavalidation.md)|_Property_ > formula2|Gets or sets the Formula2, i.e. maximum value or value depending of the operator.|1.8|
|[basicDataValidation](reference/excel/basicdatavalidation.md)|_Relationship_ > operator|The operator to use for validating the data.|1.8|
|[chart](reference/excel/chart.md)|_Property_ > categoryLabelLevel|Returns or sets a ChartCategoryLabelLevel enumeration constant referring to|1.8|
|[chart](reference/excel/chart.md)|_Property_ > chartType|Represents the type of the chart. Possible values are: ColumnClustered, ColumnStacked, ColumnStacked100, BarClustered, BarStacked, BarStacked100, LineStacked, LineStacked100, LineMarkers, LineMarkersStacked, LineMarkersStacked100, PieOfPie, etc..|1.7|
|[chart](reference/excel/chart.md)|_Property_ > id|The unique id of chart. Read-only.|1.7|
|[chart](reference/excel/chart.md)|_Property_ > plotVisibleOnly|TrueΓö¼├íif only visible cells are plotted.Γö¼├íFalseΓö¼├íif both visible and hidden cells are plotted. ReadWrite.|1.8|
|[chart](reference/excel/chart.md)|_Property_ > seriesNameLevel|Returns or sets a ChartSeriesNameLevel enumeration constant referring to|1.8|
|[chart](reference/excel/chart.md)|_Property_ > showAllFieldButtons|Represents whether to display all field buttons on a PivotChart.|1.7|
|[chart](reference/excel/chart.md)|_Property_ > showDataLabelsOverMaximum|Represents whether to to show the data labels when the value is greater than the maximum value on the value axis.|1.8|
|[chart](reference/excel/chart.md)|_Property_ > style|Returns or sets the chart style for the chart. ReadWrite.|1.8|
|[chart](reference/excel/chart.md)|_Relationship_ > displayBlanksAs|Returns or sets the way that blank cells are plotted on a chart. ReadWrite.|1.8|
|[chart](reference/excel/chart.md)|_Relationship_ > pivotOptions|Encapsulates the options for the pivot chart. Read-only.|beta|
|[chart](reference/excel/chart.md)|_Relationship_ > plotArea|Represents the plotArea for the chart. Read-only.|1.8|
|[chart](reference/excel/chart.md)|_Relationship_ > plotBy|Returns or sets the way columns or rows are used as data series on the chart. ReadWrite.|1.8|
|[chart](reference/excel/chart.md)|_Method_ > [activate()]((reference/excel/chart.md#activate)|Activate the chart in the Excel UI.|beta|
|[chartActivatedEventArgs](reference/excel/chartactivatedeventargs.md)|_Property_ > chartId|Gets the id of the chart that is activated.|1.8|
|[chartActivatedEventArgs](reference/excel/chartactivatedeventargs.md)|_Property_ > type|Gets the type of the event.|1.8|
|[chartActivatedEventArgs](reference/excel/chartactivatedeventargs.md)|_Property_ > worksheetId|Gets the id of the worksheet in which the chart is activated.|1.8|
|[chartAddedEventArgs](reference/excel/chartaddedeventargs.md)|_Property_ > chartId|Gets the id of the chart that is added to the worksheet.|1.8|
|[chartAddedEventArgs](reference/excel/chartaddedeventargs.md)|_Property_ > type|Gets the type of the event.|1.8|
|[chartAddedEventArgs](reference/excel/chartaddedeventargs.md)|_Property_ > worksheetId|Gets the id of the worksheet in which the chart is added.|1.8|
|[chartAddedEventArgs](reference/excel/chartaddedeventargs.md)|_Relationship_ > source|Gets the source of the event.|1.8|
|[chartAreaFormat](reference/excel/chartareaformat.md)|_Property_ > roundedCorners|TrueΓö¼├íif the chart area of the chart has rounded corners. ReadWrite.|beta|
|[chartAreaFormat](reference/excel/chartareaformat.md)|_Relationship_ > border|Represents the border format of chart area, which includes color, linestyle, and weight. Read-only.|1.7|
|[chartAreaFormat](reference/excel/chartareaformat.md)|_Relationship_ > colorScheme|Returns or sets anΓö¼├íintegerΓö¼├íthat represents the color scheme for the chart. ReadWrite.|beta|
|[chartAxes](reference/excel/chartaxes.md)|_Method_ > [getItem(type: string, group: string)]((reference/excel/chartaxes.md#getitemtype-string-group-string)|Returns the specific axis identified by type and group.|1.7|
|[chartAxis](reference/excel/chartaxis.md)|_Property_ > crossesAt|Represents the specified axis where the other axis crosses at. Read Only. Set to this property should use SetCrossesAt(double) method. Read-only.|1.7|
|[chartAxis](reference/excel/chartaxis.md)|_Property_ > customDisplayUnit|Represents the custom axis display unit value. Read-only. To set this property, please use the SetCustomDisplayUnit(double) method. Read-only.|1.7|
|[chartAxis](reference/excel/chartaxis.md)|_Property_ > height|Represents the height, in points, of the chart axis. Null if the axis is not visible. Read-only.|1.7|
|[chartAxis](reference/excel/chartaxis.md)|_Property_ > isBetweenCategories|Represents whether value axis crosses the category axis between categories.|1.8|
|[chartAxis](reference/excel/chartaxis.md)|_Property_ > left|Represents the distance, in points, from the left edge of the axis to the left of chart area. Null if the axis is not visible. Read-only.|1.7|
|[chartAxis](reference/excel/chartaxis.md)|_Property_ > linkNumberFormat|Represents whether the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).|beta|
|[chartAxis](reference/excel/chartaxis.md)|_Property_ > logBase|Represents the base of the logarithm when using logarithmic scales.|1.7|
|[chartAxis](reference/excel/chartaxis.md)|_Property_ > multiLevel|Represents whether an axis is multilevel or not.|1.8|
|[chartAxis](reference/excel/chartaxis.md)|_Property_ > numberFormat|Represents the format code for the axis tick label.|1.8|
|[chartAxis](reference/excel/chartaxis.md)|_Property_ > offset|Represents the distance between the levels of labels, and the distance between the first level and the axis line. The value should be an integer from 0 to 1000.|1.8|
|[chartAxis](reference/excel/chartaxis.md)|_Property_ > positionAt|Represents the specified axis position where the other axis crosses at. Read Only. Set to this property should use SetPositionAt(double) method. Read-only.|1.8|
|[chartAxis](reference/excel/chartaxis.md)|_Property_ > reversePlotOrder|Represents whether Microsoft Excel plots data points from last to first.|1.7|
|[chartAxis](reference/excel/chartaxis.md)|_Property_ > showDisplayUnitLabel|Represents whether the axis display unit label is visible.|1.7|
|[chartAxis](reference/excel/chartaxis.md)|_Property_ > textOrientation|Represents the text orientation of the axis tick label. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|1.8|
|[chartAxis](reference/excel/chartaxis.md)|_Property_ > tickLabelSpacing|Represents the number of categories or series between tick-mark labels. Can be a value from 1 through 31999 or an empty string for automatic setting. The returned value is always a number.|1.7|
|[chartAxis](reference/excel/chartaxis.md)|_Property_ > tickMarkSpacing|Represents the number of categories or series between tick marks.|1.7|
|[chartAxis](reference/excel/chartaxis.md)|_Property_ > top|Represents the distance, in points, from the top edge of the axis to the top of chart area. Null if the axis is not visible. Read-only.|1.7|
|[chartAxis](reference/excel/chartaxis.md)|_Property_ > visible|A boolean value represents the visibility of the axis.|1.7|
|[chartAxis](reference/excel/chartaxis.md)|_Property_ > width|Represents the width, in points, of the chart axis. Null if the axis is not visible. Read-only.|1.7|
|[chartAxis](reference/excel/chartaxis.md)|_Relationship_ > alignment|Represents the alignment for the specified axis tick label.|1.8|
|[chartAxis](reference/excel/chartaxis.md)|_Relationship_ > axisGroup|Represents the group for the specified axis. Read-only.|1.7|
|[chartAxis](reference/excel/chartaxis.md)|_Relationship_ > baseTimeUnit|Returns or sets the base unit for the specified category axis.|1.7|
|[chartAxis](reference/excel/chartaxis.md)|_Relationship_ > categoryType|Returns or sets the category axis type.|1.7|
|[chartAxis](reference/excel/chartaxis.md)|_Relationship_ > crosses|Represents the specified axis where the other axis crosses.|1.7|
|[chartAxis](reference/excel/chartaxis.md)|_Relationship_ > displayUnit|Represents the axis display unit.|1.7|
|[chartAxis](reference/excel/chartaxis.md)|_Relationship_ > majorTickMark|Represents the type of major tick mark for the specified axis.|1.7|
|[chartAxis](reference/excel/chartaxis.md)|_Relationship_ > majorTimeUnitScale|Returns or sets the major unit scale value for the category axis when the CategoryType property is set to TimeScale.|1.7|
|[chartAxis](reference/excel/chartaxis.md)|_Relationship_ > minorTickMark|Represents the type of minor tick mark for the specified axis.|1.7|
|[chartAxis](reference/excel/chartaxis.md)|_Relationship_ > minorTimeUnitScale|Returns or sets the minor unit scale value for the category axis when the CategoryType property is set to TimeScale.|1.7|
|[chartAxis](reference/excel/chartaxis.md)|_Relationship_ > position|Represents the specified axis position where the other axis crosses.|1.8|
|[chartAxis](reference/excel/chartaxis.md)|_Relationship_ > scaleType|Represents the value axis scale type.|1.7|
|[chartAxis](reference/excel/chartaxis.md)|_Relationship_ > tickLabelPosition|Represents the position of tick-mark labels on the specified axis.|1.7|
|[chartAxis](reference/excel/chartaxis.md)|_Relationship_ > type|Represents the axis type. Read-only.|1.7|
|[chartAxis](reference/excel/chartaxis.md)|_Method_ > [setCategoryNames(sourceData: Range)]((reference/excel/chartaxis.md#setcategorynamessourcedata-range)|Sets all the category names for the specified axis.|1.7|
|[chartAxis](reference/excel/chartaxis.md)|_Method_ > [setCrossesAt(value: double)]((reference/excel/chartaxis.md#setcrossesatvalue-double)|Set the specified axis where the other axis crosses at.|1.7|
|[chartAxis](reference/excel/chartaxis.md)|_Method_ > [setCustomDisplayUnit(value: double)]((reference/excel/chartaxis.md#setcustomdisplayunitvalue-double)|Sets the axis display unit to a custom value.|1.7|
|[chartAxis](reference/excel/chartaxis.md)|_Method_ > [setPositionAt(value: double)]((reference/excel/chartaxis.md#setpositionatvalue-double)|Set the specified axis position where the other axis crosses at.|1.8|
|[chartAxisFormat](reference/excel/chartaxisformat.md)|_Relationship_ > fill|Represents chart fill formatting. Read-only.|1.8|
|[chartAxisTitle](reference/excel/chartaxistitle.md)|_Method_ > [setFormula(formula: string)]((reference/excel/chartaxistitle.md#setformulaformula-string)|A string value that represents the formula of chart axis title using A1-style notation.|1.8|
|[chartAxisTitleFormat](reference/excel/chartaxistitleformat.md)|_Relationship_ > border|Represents the border format, which includes color, linestyle, and weight. Read-only.|1.8|
|[chartAxisTitleFormat](reference/excel/chartaxistitleformat.md)|_Relationship_ > fill|Represents chart fill formatting. Read-only.|1.8|
|[chartBinOptions](reference/excel/chartbinoptions.md)|_Property_ > allowOverflow|Returns or sets if bin overflow enabled in a histogram chart or pareto chart. ReadWrite.|beta|
|[chartBinOptions](reference/excel/chartbinoptions.md)|_Property_ > allowUnderflow|Returns or sets if bin underflow enabled in a histogram chart or pareto chart. ReadWrite.|beta|
|[chartBinOptions](reference/excel/chartbinoptions.md)|_Property_ > count|Returns or sets count of bin of a histogram chart or pareto chart. ReadWrite.|beta|
|[chartBinOptions](reference/excel/chartbinoptions.md)|_Property_ > overflowValue|Returns or sets bin overflow value of a histogram chart or pareto chart. ReadWrite.|beta|
|[chartBinOptions](reference/excel/chartbinoptions.md)|_Property_ > underflowValue|Returns or sets bin underflow value of a histogram chart or pareto chart. ReadWrite.|beta|
|[chartBinOptions](reference/excel/chartbinoptions.md)|_Property_ > width|Returns or sets bin width value of a histogram chart or pareto chart. ReadWrite.|beta|
|[chartBinOptions](reference/excel/chartbinoptions.md)|_Relationship_ > type|Returns or sets bin type of a histogram chart or pareto chart. ReadWrite.|beta|
|[chartBorder](reference/excel/chartborder.md)|_Property_ > color|HTML color code representing the color of borders in the chart.|1.7|
|[chartBorder](reference/excel/chartborder.md)|_Property_ > weight|Represents weight of the border, in points.|1.7|
|[chartBorder](reference/excel/chartborder.md)|_Relationship_ > lineStyle|Represents the line style of the border.|1.7|
|[chartBorder](reference/excel/chartborder.md)|_Method_ > [clear()]((reference/excel/chartborder.md#clear)|Clear the border format of a chart element.|1.8|
|[chartBoxwhiskerOptions](reference/excel/chartboxwhiskeroptions.md)|_Property_ > showInnerPoints|Returns or sets if inner points showed in a Box &amp; whisker chart. ReadWrite.|beta|
|[chartBoxwhiskerOptions](reference/excel/chartboxwhiskeroptions.md)|_Property_ > showMeanLine|Returns or sets if mean line showed in a Box &amp; whisker chart. ReadWrite.|beta|
|[chartBoxwhiskerOptions](reference/excel/chartboxwhiskeroptions.md)|_Property_ > showMeanMarker|Returns or sets if mean marker showed in a Box &amp; whisker chart. ReadWrite.|beta|
|[chartBoxwhiskerOptions](reference/excel/chartboxwhiskeroptions.md)|_Property_ > showOutlierPoints|Returns or sets if outlier points showed in a Box &amp; whisker chart. ReadWrite.|beta|
|[chartBoxwhiskerOptions](reference/excel/chartboxwhiskeroptions.md)|_Relationship_ > quartileCalculation|Returns or sets quartile calculation type of a Box &amp; whisker chart. ReadWrite.|beta|
|[chartDataLabel](reference/excel/chartdatalabel.md)|_Property_ > autoText|Boolean value representing if data label automatically generates appropriate text based on context.|1.8|
|[chartDataLabel](reference/excel/chartdatalabel.md)|_Property_ > formula|String value that represents the formula of chart data label using A1-style notation.|1.8|
|[chartDataLabel](reference/excel/chartdatalabel.md)|_Property_ > height|Returns the height, in points, of the chart data label. Read-only. Null if chart data label is not visible. Read-only.|1.8|
|[chartDataLabel](reference/excel/chartdatalabel.md)|_Property_ > left|Represents the distance, in points, from the left edge of chart data label to the left edge of chart area. Null if chart data label is not visible.|1.8|
|[chartDataLabel](reference/excel/chartdatalabel.md)|_Property_ > linkNumberFormat|Boolean value representing if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).|beta|
|[chartDataLabel](reference/excel/chartdatalabel.md)|_Property_ > numberFormat|String value that represents the format code for data label.|1.8|
|[chartDataLabel](reference/excel/chartdatalabel.md)|_Property_ > position|DataLabelPosition value that represents the position of the data label. Possible values are: None, Center, InsideEnd, InsideBase, OutsideEnd, Left, Right, Top, Bottom, BestFit, Callout.|1.7|
|[chartDataLabel](reference/excel/chartdatalabel.md)|_Property_ > separator|String representing the separator used for the data label on a chart.|1.7|
|[chartDataLabel](reference/excel/chartdatalabel.md)|_Property_ > showBubbleSize|Boolean value representing if the data label bubble size is visible or not.|1.7|
|[chartDataLabel](reference/excel/chartdatalabel.md)|_Property_ > showCategoryName|Boolean value representing if the data label category name is visible or not.|1.7|
|[chartDataLabel](reference/excel/chartdatalabel.md)|_Property_ > showLegendKey|Boolean value representing if the data label legend key is visible or not.|1.7|
|[chartDataLabel](reference/excel/chartdatalabel.md)|_Property_ > showPercentage|Boolean value representing if the data label percentage is visible or not.|1.7|
|[chartDataLabel](reference/excel/chartdatalabel.md)|_Property_ > showSeriesName|Boolean value representing if the data label series name is visible or not.|1.7|
|[chartDataLabel](reference/excel/chartdatalabel.md)|_Property_ > showValue|Boolean value representing if the data label value is visible or not.|1.7|
|[chartDataLabel](reference/excel/chartdatalabel.md)|_Property_ > text|String representing the text of the data label on a chart.|1.8|
|[chartDataLabel](reference/excel/chartdatalabel.md)|_Property_ > textOrientation|Represents the text orientation of chart data label. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|1.8|
|[chartDataLabel](reference/excel/chartdatalabel.md)|_Property_ > top|Represents the distance, in points, from the top edge of chart data label to the top of chart area. Null if chart data label is not visible.|1.8|
|[chartDataLabel](reference/excel/chartdatalabel.md)|_Property_ > width|Returns the width, in points, of the chart data label. Read-only. Null if chart data label is not visible. Read-only.|1.8|
|[chartDataLabel](reference/excel/chartdatalabel.md)|_Relationship_ > format|Represents the format of chart data label. Read-only.|1.8|
|[chartDataLabel](reference/excel/chartdatalabel.md)|_Relationship_ > horizontalAlignment|Represents the horizontal alignment for chart data label.|1.8|
|[chartDataLabel](reference/excel/chartdatalabel.md)|_Relationship_ > verticalAlignment|Represents the vertical alignment of chart data label.|1.8|
|[chartDataLabelFormat](reference/excel/chartdatalabelformat.md)|_Relationship_ > border|Represents the border format, which includes color, linestyle, and weight. Read-only.|1.8|
|[chartDataLabels](reference/excel/chartdatalabels.md)|_Property_ > autoText|Represents whether data labels automatically generates appropriate text based on context.|1.8|
|[chartDataLabels](reference/excel/chartdatalabels.md)|_Property_ > linkNumberFormat|Represents whether the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).|beta|
|[chartDataLabels](reference/excel/chartdatalabels.md)|_Property_ > numberFormat|Represents the format code for data labels.|1.8|
|[chartDataLabels](reference/excel/chartdatalabels.md)|_Property_ > textOrientation|Represents the text orientation of data labels. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|1.8|
|[chartDataLabels](reference/excel/chartdatalabels.md)|_Relationship_ > horizontalAlignment|Represents the horizontal alignment for chart data label.|1.8|
|[chartDataLabels](reference/excel/chartdatalabels.md)|_Relationship_ > verticalAlignment|Represents the vertical alignment of chart data label.|1.8|
|[chartDeactivatedEventArgs](reference/excel/chartdeactivatedeventargs.md)|_Property_ > chartId|Gets the id of the chart that is deactivated.|1.8|
|[chartDeactivatedEventArgs](reference/excel/chartdeactivatedeventargs.md)|_Property_ > type|Gets the type of the event.|1.8|
|[chartDeactivatedEventArgs](reference/excel/chartdeactivatedeventargs.md)|_Property_ > worksheetId|Gets the id of the worksheet in which the chart is deactivated.|1.8|
|[chartDeletedEventArgs](reference/excel/chartdeletedeventargs.md)|_Property_ > chartId|Gets the id of the chart that is deleted from the worksheet.|1.8|
|[chartDeletedEventArgs](reference/excel/chartdeletedeventargs.md)|_Property_ > type|Gets the type of the event.|1.8|
|[chartDeletedEventArgs](reference/excel/chartdeletedeventargs.md)|_Property_ > worksheetId|Gets the id of the worksheet in which the chart is deleted.|1.8|
|[chartDeletedEventArgs](reference/excel/chartdeletedeventargs.md)|_Relationship_ > source|Gets the source of the event.|1.8|
|[chartErrorBars](reference/excel/charterrorbars.md)|_Property_ > endStyleCap|Represents whether have the end style cap for the error bars.|beta|
|[chartErrorBars](reference/excel/charterrorbars.md)|_Property_ > visible|Represents whether shown error bars.|beta|
|[chartErrorBars](reference/excel/charterrorbars.md)|_Relationship_ > format|Represents the formatting of chart ErrorBars. Read-only.|beta|
|[chartErrorBars](reference/excel/charterrorbars.md)|_Relationship_ > include|Represents which error-bar parts to include.|beta|
|[chartErrorBars](reference/excel/charterrorbars.md)|_Relationship_ > type|Represents the range marked by error bars.|beta|
|[chartErrorBarsFormat](reference/excel/charterrorbarsformat.md)|_Relationship_ > line|Represents chart line formatting. Read-only.|beta|
|[chartFormatString](reference/excel/chartformatstring.md)|_Relationship_ > font|Represents the font attributes, such as font name, font size, color, etc. of chart characters object. Read-only.|1.7|
|[chartLegend](reference/excel/chartlegend.md)|_Property_ > height|Represents the height, in points, of the legend on the chart. Null if legend is not visible.|1.7|
|[chartLegend](reference/excel/chartlegend.md)|_Property_ > left|Represents the left, in points, of a chart legend. Null if legend is not visible.|1.7|
|[chartLegend](reference/excel/chartlegend.md)|_Property_ > showShadow|Represents if the legend has a shadow on the chart.|1.7|
|[chartLegend](reference/excel/chartlegend.md)|_Property_ > top|Represents the top of a chart legend.|1.7|
|[chartLegend](reference/excel/chartlegend.md)|_Property_ > width|Represents the width, in points, of the legend on the chart. Null if legend is not visible.|1.7|
|[chartLegend](reference/excel/chartlegend.md)|_Relationship_ > legendEntries|Represents a collection of legendEntries in the legend. Read-only.|1.7|
|[chartLegendEntry](reference/excel/chartlegendentry.md)|_Property_ > height|Represents the height of the legendEntry on the chart Legend. Read-only.|1.8|
|[chartLegendEntry](reference/excel/chartlegendentry.md)|_Property_ > index|Represents the index of the LegendEntry in the Chart Legend. Read-only.|1.8|
|[chartLegendEntry](reference/excel/chartlegendentry.md)|_Property_ > left|Represents the left of a chart legendEntry. Read-only.|1.8|
|[chartLegendEntry](reference/excel/chartlegendentry.md)|_Property_ > top|Represents the top of a chart legendEntry. Read-only.|1.8|
|[chartLegendEntry](reference/excel/chartlegendentry.md)|_Property_ > visible|Represents the visible of a chart legend entry.|1.7|
|[chartLegendEntry](reference/excel/chartlegendentry.md)|_Property_ > width|Represents the width of the legendEntry on the chart Legend. Read-only.|1.8|
|[chartLegendEntryCollection](reference/excel/chartlegendentrycollection.md)|_Property_ > items|A collection of chartLegendEntry objects. Read-only.|1.7|
|[chartLegendEntryCollection](reference/excel/chartlegendentrycollection.md)|_Method_ > [getCount()]((reference/excel/chartlegendentrycollection.md#getcount)|Returns the number of legendEntry in the collection.|1.7|
|[chartLegendEntryCollection](reference/excel/chartlegendentrycollection.md)|_Method_ > [getItemAt(index: number)]((reference/excel/chartlegendentrycollection.md#getitematindex-number)|Returns a legendEntry at the given index.|1.7|
|[chartLegendFormat](reference/excel/chartlegendformat.md)|_Relationship_ > border|Represents the border format, which includes color, linestyle, and weight. Read-only.|1.8|
|[chartLineFormat](reference/excel/chartlineformat.md)|_Property_ > weight|Represents weight of the line, in points.|1.7|
|[chartLineFormat](reference/excel/chartlineformat.md)|_Relationship_ > lineStyle|Represents the line style.|1.7|
|[chartMapOptions](reference/excel/chartmapoptions.md)|_Relationship_ > labelStrategy|Returns or sets series map labels strategy of a region map chart. ReadWrite.|beta|
|[chartMapOptions](reference/excel/chartmapoptions.md)|_Relationship_ > level|Returns or sets series map area of a region map chart. ReadWrite.|beta|
|[chartMapOptions](reference/excel/chartmapoptions.md)|_Relationship_ > projectionType|Returns or sets series projection type of a region map chart. ReadWrite.|beta|
|[chartPivotOptions](reference/excel/chartpivotoptions.md)|_Property_ > showAxisFieldButtons|Represents whether to display axis field buttons on a PivotChart.|beta|
|[chartPivotOptions](reference/excel/chartpivotoptions.md)|_Property_ > showLegendFieldButtons|Represents whether to display legend field buttons on a PivotChart.|beta|
|[chartPivotOptions](reference/excel/chartpivotoptions.md)|_Property_ > showReportFilterFieldButtons|Represents whether to display report filter field buttons on a PivotChart.|beta|
|[chartPivotOptions](reference/excel/chartpivotoptions.md)|_Property_ > showValueFieldButtons|Represents whether to display show value field buttons on a PivotChart.|beta|
|[chartPlotArea](reference/excel/chartplotarea.md)|_Property_ > height|Represents the height value of plotArea.|1.8|
|[chartPlotArea](reference/excel/chartplotarea.md)|_Property_ > insideHeight|Represents the insideHeight value of plotArea.|1.8|
|[chartPlotArea](reference/excel/chartplotarea.md)|_Property_ > insideLeft|Represents the insideLeft value of plotArea.|1.8|
|[chartPlotArea](reference/excel/chartplotarea.md)|_Property_ > insideTop|Represents the insideTop value of plotArea.|1.8|
|[chartPlotArea](reference/excel/chartplotarea.md)|_Property_ > insideWidth|Represents the insideWidth value of plotArea.|1.8|
|[chartPlotArea](reference/excel/chartplotarea.md)|_Property_ > left|Represents the left value of plotArea.|1.8|
|[chartPlotArea](reference/excel/chartplotarea.md)|_Property_ > top|Represents the top value of plotArea.|1.8|
|[chartPlotArea](reference/excel/chartplotarea.md)|_Property_ > width|Represents the width value of plotArea.|1.8|
|[chartPlotArea](reference/excel/chartplotarea.md)|_Relationship_ > format|Represents the formatting of a chart plotArea. Read-only.|1.8|
|[chartPlotArea](reference/excel/chartplotarea.md)|_Relationship_ > position|Represents the position of plotArea.|1.8|
|[chartPlotAreaFormat](reference/excel/chartplotareaformat.md)|_Relationship_ > border|Represents the border attributes of a chart plotArea. Read-only.|1.8|
|[chartPlotAreaFormat](reference/excel/chartplotareaformat.md)|_Relationship_ > fill|Represents the fill format of an object, which includes background formating information. Read-only.|1.8|
|[chartPoint](reference/excel/chartpoint.md)|_Property_ > hasDataLabel|Represents whether a data point has a data label. Not applicable for surface charts.|1.7|
|[chartPoint](reference/excel/chartpoint.md)|_Property_ > markerBackgroundColor|HTML color code representation of the marker background color of data point. E.g. #FF0000 represents Red.|1.7|
|[chartPoint](reference/excel/chartpoint.md)|_Property_ > markerForegroundColor|HTML color code representation of the marker foreground color of data point. E.g. #FF0000 represents Red.|1.7|
|[chartPoint](reference/excel/chartpoint.md)|_Property_ > markerSize|Represents marker size of data point.|1.7|
|[chartPoint](reference/excel/chartpoint.md)|_Relationship_ > dataLabel|Returns the data label of a chart point. Read-only.|1.7|
|[chartPoint](reference/excel/chartpoint.md)|_Relationship_ > markerStyle|Represents marker style of a chart data point.|1.7|
|[chartPointFormat](reference/excel/chartpointformat.md)|_Relationship_ > border|Represents the border format of a chart data point, which includes color, style, and weight information. Read-only.|1.7|
|[chartPointsCollection](reference/excel/chartpointscollection.md)|_Method_ > [getFirst()]((reference/excel/chartpointscollection.md#getfirst)|Gets the first point in the series.|ApiSetAttribute.Spec|
|[chartPointsCollection](reference/excel/chartpointscollection.md)|_Method_ > [getLast()]((reference/excel/chartpointscollection.md#getlast)|Gets the last point in the series.|ApiSetAttribute.Spec|
|[chartSeries](reference/excel/chartseries.md)|_Property_ > bubbleScale|Returns or sets the scale factor for bubbles in the specified chart group. Can be an integer value from 0 (zero) to 300, corresponding to a percentage of the default size. Applies only to bubble charts. ReadWrite.|beta|
|[chartSeries](reference/excel/chartseries.md)|_Property_ > chartType|Represents the chart type of a series. Possible values are: ColumnClustered, ColumnStacked, ColumnStacked100, BarClustered, BarStacked, BarStacked100, LineStacked, LineStacked100, LineMarkers, LineMarkersStacked, LineMarkersStacked100, PieOfPie, etc..|1.7|
|[chartSeries](reference/excel/chartseries.md)|_Property_ > doughnutHoleSize|Represents the doughnut hole size of a chart series.  Only valid on doughnut and doughnutExploded charts.|1.7|
|[chartSeries](reference/excel/chartseries.md)|_Property_ > explosion|Returns or sets the explosion value for a pie-chart or doughnut-chart slice. Returns 0 (zero) if there's no explosion (the tip of the slice is in the center of the pie). ReadWrite.|1.8|
|[chartSeries](reference/excel/chartseries.md)|_Property_ > filtered|Boolean value representing if the series is filtered or not. Not applicable for surface charts.|1.7|
|[chartSeries](reference/excel/chartseries.md)|_Property_ > firstSliceAngle|Returns or sets the angle of the first pie-chart or doughnut-chart slice, in degrees (clockwise from vertical). Applies only to pie, 3-D pie, and doughnut charts. Can be a value from 0 through 360. ReadWrite|1.8|
|[chartSeries](reference/excel/chartseries.md)|_Property_ > gapWidth|Represents the gap width of a chart series.  Only valid on bar and column charts, as well as|1.7|
|[chartSeries](reference/excel/chartseries.md)|_Property_ > gradientMaximumColor|Returns or sets the Color for maximum value of a region map chart series. ReadWrite.|beta|
|[chartSeries](reference/excel/chartseries.md)|_Property_ > gradientMaximumValue|Returns or sets the maximum value of a region map chart series. ReadWrite.|beta|
|[chartSeries](reference/excel/chartseries.md)|_Property_ > gradientMidpointColor|Returns or sets the Color for midpoint value of a region map chart series. ReadWrite.|beta|
|[chartSeries](reference/excel/chartseries.md)|_Property_ > gradientMidpointValue|Returns or sets the midpoint value of a region map chart series. ReadWrite.|beta|
|[chartSeries](reference/excel/chartseries.md)|_Property_ > gradientMinimumColor|Returns or sets the Color for minimum value of a region map chart series. ReadWrite.|beta|
|[chartSeries](reference/excel/chartseries.md)|_Property_ > gradientMinimumValue|Returns or sets the minimum value of a region map chart series. ReadWrite.|beta|
|[chartSeries](reference/excel/chartseries.md)|_Property_ > hasDataLabels|Boolean value representing if the series has data labels or not.|1.7|
|[chartSeries](reference/excel/chartseries.md)|_Property_ > invertColor|Returns or sets the fill color for negative data points in a series. ReadWrite.|beta|
|[chartSeries](reference/excel/chartseries.md)|_Property_ > invertIfNegative|TrueΓö¼├íif Microsoft Excel inverts the pattern in the item when it corresponds to a negative number. ReadWrite.|1.8|
|[chartSeries](reference/excel/chartseries.md)|_Property_ > markerBackgroundColor|Represents markers background color of a chart series.|1.7|
|[chartSeries](reference/excel/chartseries.md)|_Property_ > markerForegroundColor|Represents markers foreground color of a chart series.|1.7|
|[chartSeries](reference/excel/chartseries.md)|_Property_ > markerSize|Represents marker size of a chart series.|1.7|
|[chartSeries](reference/excel/chartseries.md)|_Property_ > overlap|Specifies how bars and columns are positioned. Can be a value between ╬ô├ç├┤ 100 and 100. Applies only to 2-D bar and 2-D column charts. ReadWrite.|1.8|
|[chartSeries](reference/excel/chartseries.md)|_Property_ > plotOrder|Represents the plot order of a chart series within the chart group.|1.7|
|[chartSeries](reference/excel/chartseries.md)|_Property_ > secondPlotSize|Returns or sets the size of the secondary section of either a pie of pie chart or a bar of pie chart, as a percentage of the size of the primary pie. Can be a value from 5 to 200. ReadWrite.|1.8|
|[chartSeries](reference/excel/chartseries.md)|_Property_ > showConnectorLines|Returns or sets if connector lines show in a waterfall chart. ReadWrite.|beta|
|[chartSeries](reference/excel/chartseries.md)|_Property_ > showLeaderLines|True if Microsoft Excel show leaderlines for each datalabel in series. ReadWrite.|beta|
|[chartSeries](reference/excel/chartseries.md)|_Property_ > showShadow|Boolean value representing if the series has a shadow or not.|1.7|
|[chartSeries](reference/excel/chartseries.md)|_Property_ > smooth|Boolean value representing if the series is smooth or not. Only applicable to line and scatter charts.|1.7|
|[chartSeries](reference/excel/chartseries.md)|_Property_ > splitValue|Returns or sets the threshold value separating the two sections of either a pie of pie chart or a bar of pie chart. ReadWrite.|beta|
|[chartSeries](reference/excel/chartseries.md)|_Property_ > varyByCategories|TrueΓö¼├íif Microsoft Excel assigns a different color or pattern to each data marker. The chart must contain only one series. ReadWrite.|1.8|
|[chartSeries](reference/excel/chartseries.md)|_Relationship_ > axisGroup|Returns or sets the group for the specified series. ReadWrite|1.8|
|[chartSeries](reference/excel/chartseries.md)|_Relationship_ > binOptions|Encapsulates the bin options only for histogram chart and pareto chart. Read-only.|beta|
|[chartSeries](reference/excel/chartseries.md)|_Relationship_ > boxwhiskerOptions|Encapsulates the options for the Box &amp; Whisker chart. Read-only.|beta|
|[chartSeries](reference/excel/chartseries.md)|_Relationship_ > dataLabels|Represents a collection of all dataLabels in the series. Read-only.|1.8|
|[chartSeries](reference/excel/chartseries.md)|_Relationship_ > gradientMaximumType|Returns or sets the type for maximum value of a region map chart series. ReadWrite.|beta|
|[chartSeries](reference/excel/chartseries.md)|_Relationship_ > gradientMidpointType|Returns or sets the type for midpoint value of a region map chart series. ReadWrite.|beta|
|[chartSeries](reference/excel/chartseries.md)|_Relationship_ > gradientMinimumType|Returns or sets the type for minimum value of a region map chart series. ReadWrite.|beta|
|[chartSeries](reference/excel/chartseries.md)|_Relationship_ > gradientStyle|Returns or sets series gradient style of a region map chart. ReadWrite.|beta|
|[chartSeries](reference/excel/chartseries.md)|_Relationship_ > mapOptions|Encapsulates the options for the Map chart. Read-only.|beta|
|[chartSeries](reference/excel/chartseries.md)|_Relationship_ > markerStyle|Represents marker style of a chart series.|1.7|
|[chartSeries](reference/excel/chartseries.md)|_Relationship_ > parentLabelStrategy|Returns or sets series parent label strategy area of a treemap chart. ReadWrite.|beta|
|[chartSeries](reference/excel/chartseries.md)|_Relationship_ > splitType|Returns or sets the way the two sections of either a pie of pie chart or a bar of pie chart are split. ReadWrite.|1.8|
|[chartSeries](reference/excel/chartseries.md)|_Relationship_ > trendlines|Represents a collection of trendlines in the series. Read-only.|1.7|
|[chartSeries](reference/excel/chartseries.md)|_Relationship_ > xErrorBars|Represents the error bar object for a chart series. Read-only.|beta|
|[chartSeries](reference/excel/chartseries.md)|_Relationship_ > yErrorBars|Represents the error bar object for a chart series. Read-only.|beta|
|[chartSeries](reference/excel/chartseries.md)|_Method_ > [delete()]((reference/excel/chartseries.md#delete)|Deletes the chart series.|1.7|
|[chartSeries](reference/excel/chartseries.md)|_Method_ > [setBubbleSizes(sourceData: Range)]((reference/excel/chartseries.md#setbubblesizessourcedata-range)|Set bubble sizes for a chart series. Only works for bubble charts.|1.7|
|[chartSeries](reference/excel/chartseries.md)|_Method_ > [setValues(sourceData: Range)]((reference/excel/chartseries.md#setvaluessourcedata-range)|Set values for a chart series. For scatter chart, it means Y axis values.|1.7|
|[chartSeries](reference/excel/chartseries.md)|_Method_ > [setXAxisValues(sourceData: Range)]((reference/excel/chartseries.md#setxaxisvaluessourcedata-range)|Set values of X axis for a chart series. Only works for scatter charts.|1.7|
|[chartSeriesCollection](reference/excel/chartseriescollection.md)|_Method_ > [add(name: string, index: number)]((reference/excel/chartseriescollection.md#addname-string-index-number)|Add a new series to the collection. The new added series is not visible until set valuesx axis valuesbubble sizes for it (depending on chart type).|1.7|
|[chartSeriesCollection](reference/excel/chartseriescollection.md)|_Method_ > [getFirst()]((reference/excel/chartseriescollection.md#getfirst)|Gets the first series in the collection.|ApiSetAttribute.Spec|
|[chartSeriesCollection](reference/excel/chartseriescollection.md)|_Method_ > [getLast()]((reference/excel/chartseriescollection.md#getlast)|Gets the last series in the collection.|ApiSetAttribute.Spec|
|[chartTitle](reference/excel/charttitle.md)|_Property_ > height|Returns the height, in points, of the chart title. Null if chart title is not visible. Read-only.|1.7|
|[chartTitle](reference/excel/charttitle.md)|_Property_ > left|Represents the distance, in points, from the left edge of chart title to the left edge of chart area. Null if chart title is not visible.|1.7|
|[chartTitle](reference/excel/charttitle.md)|_Property_ > showShadow|Represents a boolean value that determines if the chart title has a shadow.|1.7|
|[chartTitle](reference/excel/charttitle.md)|_Property_ > textOrientation|Represents the text orientation of chart title. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|1.7|
|[chartTitle](reference/excel/charttitle.md)|_Property_ > top|Represents the distance, in points, from the top edge of chart title to the top of chart area. Null if chart title is not visible.|1.7|
|[chartTitle](reference/excel/charttitle.md)|_Property_ > width|Returns the width, in points, of the chart title. Null if chart title is not visible. Read-only.|1.7|
|[chartTitle](reference/excel/charttitle.md)|_Relationship_ > horizontalAlignment|Represents the horizontal alignment for chart title.|1.7|
|[chartTitle](reference/excel/charttitle.md)|_Relationship_ > position|Represents the position of chart title.|1.7|
|[chartTitle](reference/excel/charttitle.md)|_Relationship_ > verticalAlignment|Represents the vertical alignment of chart title.|1.7|
|[chartTitle](reference/excel/charttitle.md)|_Method_ > [getSubstring(start: number, length: number)]((reference/excel/charttitle.md#getsubstringstart-number-length-number)|Get the substring of a chart title. Line break '\n' also counts one character.|1.7|
|[chartTitle](reference/excel/charttitle.md)|_Method_ > [setFormula(formula: string)]((reference/excel/charttitle.md#setformulaformula-string)|Sets a string value that represents the formula of chart title using A1-style notation.|1.7|
|[chartTitleFormat](reference/excel/charttitleformat.md)|_Relationship_ > border|Represents the border format of chart title, which includes color, linestyle, and weight. Read-only.|1.7|
|[chartTrendline](reference/excel/charttrendline.md)|_Property_ > backwardPeriod|Represents the number of periods that the trendline extends backward.|1.8|
|[chartTrendline](reference/excel/charttrendline.md)|_Property_ > forwardPeriod|Represents the number of periods that the trendline extends forward.|1.8|
|[chartTrendline](reference/excel/charttrendline.md)|_Property_ > intercept|Represents the intercept value of the trendline. Can be set to a numeric value or an empty string (for automatic values). The returned value is always a number.|1.7|
|[chartTrendline](reference/excel/charttrendline.md)|_Property_ > movingAveragePeriod|Represents the period of a chart trendline. Only applicable for trendline with MovingAverage type.|1.7|
|[chartTrendline](reference/excel/charttrendline.md)|_Property_ > name|Represents the name of the trendline. Can be set to a string value, or can be set to null value represents automatic values. The returned value is always a string|1.7|
|[chartTrendline](reference/excel/charttrendline.md)|_Property_ > polynomialOrder|Represents the order of a chart trendline. Only applicable for trendline with Polynomial type.|1.7|
|[chartTrendline](reference/excel/charttrendline.md)|_Property_ > showEquation|True if the equation for the trendline is displayed on the chart.|1.8|
|[chartTrendline](reference/excel/charttrendline.md)|_Property_ > showRSquared|True if the R-squared for the trendline is displayed on the chart.|1.8|
|[chartTrendline](reference/excel/charttrendline.md)|_Relationship_ > format|Represents the formatting of a chart trendline. Read-only.|1.7|
|[chartTrendline](reference/excel/charttrendline.md)|_Relationship_ > label|Represents the label of a chart trendline. Read-only.|1.8|
|[chartTrendline](reference/excel/charttrendline.md)|_Relationship_ > type|Represents the type of a chart trendline.|1.7|
|[chartTrendline](reference/excel/charttrendline.md)|_Method_ > [delete()]((reference/excel/charttrendline.md#delete)|Delete the trendline object.|1.7|
|[chartTrendlineCollection](reference/excel/charttrendlinecollection.md)|_Property_ > items|A collection of chartTrendline objects. Read-only.|1.7|
|[chartTrendlineCollection](reference/excel/charttrendlinecollection.md)|_Method_ > [add(type: string)]((reference/excel/charttrendlinecollection.md#addtype-string)|Adds a new trendline to trendline collection.|1.7|
|[chartTrendlineCollection](reference/excel/charttrendlinecollection.md)|_Method_ > [getCount()]((reference/excel/charttrendlinecollection.md#getcount)|Returns the number of trendlines in the collection.|1.7|
|[chartTrendlineCollection](reference/excel/charttrendlinecollection.md)|_Method_ > [getItem(index: number)]((reference/excel/charttrendlinecollection.md#getitemindex-number)|Get trendline object by index, which is the insertion order in items array.|1.7|
|[chartTrendlineFormat](reference/excel/charttrendlineformat.md)|_Relationship_ > line|Represents chart line formatting. Read-only.|1.7|
|[chartTrendlineLabel](reference/excel/charttrendlinelabel.md)|_Property_ > autoText|Boolean value representing if trendline label automatically generates appropriate text based on context.|1.8|
|[chartTrendlineLabel](reference/excel/charttrendlinelabel.md)|_Property_ > formula|String value that represents the formula of chart trendline label using A1-style notation.|1.8|
|[chartTrendlineLabel](reference/excel/charttrendlinelabel.md)|_Property_ > height|Returns the height, in points, of the chart trendline label. Read-only. Null if chart trendline label is not visible. Read-only.|1.8|
|[chartTrendlineLabel](reference/excel/charttrendlinelabel.md)|_Property_ > left|Represents the distance, in points, from the left edge of chart trendline label to the left edge of chart area. Null if chart trendline label is not visible.|1.8|
|[chartTrendlineLabel](reference/excel/charttrendlinelabel.md)|_Property_ > linkNumberFormat|Boolean value representing if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).|beta|
|[chartTrendlineLabel](reference/excel/charttrendlinelabel.md)|_Property_ > numberFormat|String value that represents the format code for trendline label.|1.8|
|[chartTrendlineLabel](reference/excel/charttrendlinelabel.md)|_Property_ > text|String representing the text of the trendline label on a chart.|1.8|
|[chartTrendlineLabel](reference/excel/charttrendlinelabel.md)|_Property_ > textOrientation|Represents the text orientation of chart trendline label. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|1.8|
|[chartTrendlineLabel](reference/excel/charttrendlinelabel.md)|_Property_ > top|Represents the distance, in points, from the top edge of chart trendline label to the top of chart area. Null if chart trendline label is not visible.|1.8|
|[chartTrendlineLabel](reference/excel/charttrendlinelabel.md)|_Property_ > width|Returns the width, in points, of the chart trendline label. Read-only. Null if chart trendline label is not visible. Read-only.|1.8|
|[chartTrendlineLabel](reference/excel/charttrendlinelabel.md)|_Relationship_ > format|Represents the format of chart trendline label. Read-only.|1.8|
|[chartTrendlineLabel](reference/excel/charttrendlinelabel.md)|_Relationship_ > horizontalAlignment|Represents the horizontal alignment for chart trendline label.|1.8|
|[chartTrendlineLabel](reference/excel/charttrendlinelabel.md)|_Relationship_ > verticalAlignment|Represents the vertical alignment of chart trendline label.|1.8|
|[chartTrendlineLabelFormat](reference/excel/charttrendlinelabelformat.md)|_Relationship_ > border|Represents the border format, which includes color, linestyle, and weight. Read-only.|1.8|
|[chartTrendlineLabelFormat](reference/excel/charttrendlinelabelformat.md)|_Relationship_ > fill|Represents the fill format of the current chart trendline label. Read-only.|1.8|
|[chartTrendlineLabelFormat](reference/excel/charttrendlinelabelformat.md)|_Relationship_ > font|Represents the font attributes (font name, font size, color, etc.) for a chart trendline label. Read-only.|1.8|
|[closeWorkbookPostProcessAction](reference/excel/closeworkbookpostprocessaction.md)|_Relationship_ > actionType|Transmits additional data to client side, e.g., worksheetId for TableSelectionChangedEvent.|beta|
|[conditionalFormat](reference/excel/conditionalformat.md)|_Method_ > [getRanges()]((reference/excel/conditionalformat.md#getranges)|Returns the RangeAreas, comprising one or more rectangular ranges, the conditonal format is applied to. Read-only.|beta|
|[createWorkbookPostProcessAction](reference/excel/createworkbookpostprocessaction.md)|_Property_ > fakeFileId|Transmits additional data to client side, e.g., worksheetId for TableSelectionChangedEvent.|1.8|
|[createWorkbookPostProcessAction](reference/excel/createworkbookpostprocessaction.md)|_Property_ > fileBase64|Transmits additional data to client side, e.g., worksheetId for TableSelectionChangedEvent.|1.8|
|[createWorkbookPostProcessAction](reference/excel/createworkbookpostprocessaction.md)|_Relationship_ > actionType|Transmits additional data to client side, e.g., worksheetId for TableSelectionChangedEvent.|1.8|
|[customDataValidation](reference/excel/customdatavalidation.md)|_Property_ > formula|Custom data validation formula, it is to create special rules, such as preventing duplicates, or limiting the total in a range of cells.|1.8|
|[customProperty](reference/excel/customproperty.md)|_Property_ > key|Gets the key of the custom property. Read only. Read-only.|1.7|
|[customProperty](reference/excel/customproperty.md)|_Property_ > value|Gets or sets the value of the custom property.|1.7|
|[customProperty](reference/excel/customproperty.md)|_Relationship_ > type|Gets the value type of the custom property. Read only. Read-only.|1.7|
|[customProperty](reference/excel/customproperty.md)|_Method_ > [delete()]((reference/excel/customproperty.md#delete)|Deletes the custom property.|1.7|
|[customPropertyCollection](reference/excel/custompropertycollection.md)|_Property_ > items|A collection of customProperty objects. Read-only.|1.7|
|[customPropertyCollection](reference/excel/custompropertycollection.md)|_Method_ > [add(key: string, value: object)]((reference/excel/custompropertycollection.md#addkey-string-value-object)|Creates a new or sets an existing custom property.|1.7|
|[customPropertyCollection](reference/excel/custompropertycollection.md)|_Method_ > [deleteAll()]((reference/excel/custompropertycollection.md#deleteall)|Deletes all custom properties in this collection.|1.7|
|[customPropertyCollection](reference/excel/custompropertycollection.md)|_Method_ > [getCount()]((reference/excel/custompropertycollection.md#getcount)|Gets the count of custom properties.|1.7|
|[customPropertyCollection](reference/excel/custompropertycollection.md)|_Method_ > [getItem(key: string)]((reference/excel/custompropertycollection.md#getitemkey-string)|Gets a custom property object by its key, which is case-insensitive. Throws if the custom property does not exist.|1.7|
|[customPropertyCollection](reference/excel/custompropertycollection.md)|_Method_ > [getItemOrNullObject(key: string)]((reference/excel/custompropertycollection.md#getitemornullobjectkey-string)|Gets a custom property object by its key, which is case-insensitive. Returns a null object if the custom property does not exist.|1.7|
|[customXmlPart](reference/excel/customxmlpart.md)|_Method_ > [deleteAttribute(xpath: string, namespaceMappings: object, name: string)]((reference/excel/customxmlpart.md#deleteattributexpath-string-namespacemappings-object-name-string)|Deletes an attribute with the given name from the element identified by xpath.|ApiSetAttribute.Spec|
|[customXmlPart](reference/excel/customxmlpart.md)|_Method_ > [deleteElement(xpath: string, namespaceMappings: object)]((reference/excel/customxmlpart.md#deleteelementxpath-string-namespacemappings-object)|Deletes the element identified by xpath.|ApiSetAttribute.Spec|
|[customXmlPart](reference/excel/customxmlpart.md)|_Method_ > [insertAttribute(xpath: string, namespaceMappings: object, name: string, value: string)]((reference/excel/customxmlpart.md#insertattributexpath-string-namespacemappings-object-name-string-value-string)|Inserts an attribute with the given name and value to the element identified by xpath.|ApiSetAttribute.Spec|
|[customXmlPart](reference/excel/customxmlpart.md)|_Method_ > [insertElement(xpath: string, xml: string, namespaceMappings: object, index: number)]((reference/excel/customxmlpart.md#insertelementxpath-string-xml-string-namespacemappings-object-index-number)|Inserts the given XML under the parent element identified by xpath at child position index.|ApiSetAttribute.Spec|
|[customXmlPart](reference/excel/customxmlpart.md)|_Method_ > [query(xpath: string, namespaceMappings: object)]((reference/excel/customxmlpart.md#queryxpath-string-namespacemappings-object)|Queries the XML content.|ApiSetAttribute.Spec|
|[customXmlPart](reference/excel/customxmlpart.md)|_Method_ > [updateAttribute(xpath: string, namespaceMappings: object, name: string, value: string)]((reference/excel/customxmlpart.md#updateattributexpath-string-namespacemappings-object-name-string-value-string)|Updates the value of an attribute with the given name of the element identified by xpath.|ApiSetAttribute.Spec|
|[customXmlPart](reference/excel/customxmlpart.md)|_Method_ > [updateElement(xpath: string, xml: string, namespaceMappings: object)]((reference/excel/customxmlpart.md#updateelementxpath-string-xml-string-namespacemappings-object)|Updates the XML of the element identified by xpath.|ApiSetAttribute.Spec|
|[dataPivotHierarchy](reference/excel/datapivothierarchy.md)|_Property_ > id|Id of the DataPivotHierarchy. Read-only.|1.8|
|[dataPivotHierarchy](reference/excel/datapivothierarchy.md)|_Property_ > name|Name of the DataPivotHierarchy.|1.8|
|[dataPivotHierarchy](reference/excel/datapivothierarchy.md)|_Property_ > numberFormat|Number format of the DataPivotHierarchy.|1.8|
|[dataPivotHierarchy](reference/excel/datapivothierarchy.md)|_Property_ > position|Position of the DataPivotHierarchy.|1.8|
|[dataPivotHierarchy](reference/excel/datapivothierarchy.md)|_Relationship_ > field|Returns the PivotFields associated with the DataPivotHierarchy. Read-only.|1.8|
|[dataPivotHierarchy](reference/excel/datapivothierarchy.md)|_Relationship_ > showAs|Determines whether the data should be sown as a specific summary calculation or not.|1.8|
|[dataPivotHierarchy](reference/excel/datapivothierarchy.md)|_Relationship_ > summarizeBy|Determines whether to show all items of the DataPivotHierarchy.|1.8|
|[dataPivotHierarchy](reference/excel/datapivothierarchy.md)|_Method_ > [setToDefault()]((reference/excel/datapivothierarchy.md#settodefault)|Reset the DataPivotHierarchy back to its default values.|1.8|
|[dataPivotHierarchyCollection](reference/excel/datapivothierarchycollection.md)|_Property_ > items|A collection of dataPivotHierarchy objects. Read-only.|1.8|
|[dataPivotHierarchyCollection](reference/excel/datapivothierarchycollection.md)|_Method_ > [add(pivotHierarchy: PivotHierarchy)]((reference/excel/datapivothierarchycollection.md#addpivothierarchy-pivothierarchy)|Adds the PivotHierarchy to the current axis.|1.8|
|[dataPivotHierarchyCollection](reference/excel/datapivothierarchycollection.md)|_Method_ > [getCount()]((reference/excel/datapivothierarchycollection.md#getcount)|Gets the number of pivot hierarchies in the collection.|1.8|
|[dataPivotHierarchyCollection](reference/excel/datapivothierarchycollection.md)|_Method_ > [getItem(name: string)]((reference/excel/datapivothierarchycollection.md#getitemname-string)|Gets a DataPivotHierarchy by its name or id.|1.8|
|[dataPivotHierarchyCollection](reference/excel/datapivothierarchycollection.md)|_Method_ > [getItemOrNullObject(name: string)]((reference/excel/datapivothierarchycollection.md#getitemornullobjectname-string)|Gets a DataPivotHierarchy by name. If the DataPivotHierarchy does not exist, will return a null object.|1.8|
|[dataPivotHierarchyCollection](reference/excel/datapivothierarchycollection.md)|_Method_ > [remove(DataPivotHierarchy: DataPivotHierarchy)]((reference/excel/datapivothierarchycollection.md#removedatapivothierarchy-datapivothierarchy)|Removes the PivotHierarchy from the current axis.|1.8|
|[dataValidation](reference/excel/datavalidation.md)|_Property_ > ignoreBlanks|Ignore blanks: no data validation will be performed on blank cells, it defaults to true.|1.8|
|[dataValidation](reference/excel/datavalidation.md)|_Property_ > valid|Represents if all cell values are valid according to the data validation rules. Read-only.|1.8|
|[dataValidation](reference/excel/datavalidation.md)|_Relationship_ > errorAlert|Error alert when user enters invalid data.|1.8|
|[dataValidation](reference/excel/datavalidation.md)|_Relationship_ > prompt|Prompt when users select a cell.|1.8|
|[dataValidation](reference/excel/datavalidation.md)|_Relationship_ > rule|Data Validation rule that contains different type of data validation criteria.|1.8|
|[dataValidation](reference/excel/datavalidation.md)|_Relationship_ > type|Type of the data validation, see Excel.DataValidationType for details. Read-only.|1.8|
|[dataValidation](reference/excel/datavalidation.md)|_Method_ > [clear()]((reference/excel/datavalidation.md#clear)|Clears the data validation from the current range.|1.8|
|[dataValidation](reference/excel/datavalidation.md)|_Method_ > [getInvalidCells()]((reference/excel/datavalidation.md#getinvalidcells)|Returns a RangeAreas, comprising one or more rectangular ranges, with invalid cell values. If all cell values are valid, this function will throw an ItemNotFound error.|beta|
|[dataValidation](reference/excel/datavalidation.md)|_Method_ > [getInvalidCellsOrNullObject()]((reference/excel/datavalidation.md#getinvalidcellsornullobject)|Returns a RangeAreas, comprising one or more rectangular ranges, with invalid cell values. If all cell values are valid, this function will return null.|beta|
|[dataValidationErrorAlert](reference/excel/datavalidationerroralert.md)|_Property_ > message|Represents error alert message.|1.8|
|[dataValidationErrorAlert](reference/excel/datavalidationerroralert.md)|_Property_ > showAlert|It determines show error alert dialog or not when users enter invalid data, it defaults to true.|1.8|
|[dataValidationErrorAlert](reference/excel/datavalidationerroralert.md)|_Property_ > title|Represents error alert dialog title.|1.8|
|[dataValidationErrorAlert](reference/excel/datavalidationerroralert.md)|_Relationship_ > style|Represents Data validation alert type, please see Excel.DataValidationAlertStyle for details.|1.8|
|[dataValidationPrompt](reference/excel/datavalidationprompt.md)|_Property_ > message|Represents the message of the prompt.|1.8|
|[dataValidationPrompt](reference/excel/datavalidationprompt.md)|_Property_ > showPrompt|It determines showing the prompt or not when user selects a cell with the data validation.|1.8|
|[dataValidationPrompt](reference/excel/datavalidationprompt.md)|_Property_ > title|Represents the title for the prompt.|1.8|
|[dataValidationRule](reference/excel/datavalidationrule.md)|_Relationship_ > custom|Custom data validation criteria.|1.8|
|[dataValidationRule](reference/excel/datavalidationrule.md)|_Relationship_ > date|Date data validation criteria.|1.8|
|[dataValidationRule](reference/excel/datavalidationrule.md)|_Relationship_ > decimal|Decimal data validation criteria.|1.8|
|[dataValidationRule](reference/excel/datavalidationrule.md)|_Relationship_ > list|List data validation criteria.|1.8|
|[dataValidationRule](reference/excel/datavalidationrule.md)|_Relationship_ > textLength|TextLength data validation criteria.|1.8|
|[dataValidationRule](reference/excel/datavalidationrule.md)|_Relationship_ > time|Time data validation criteria.|1.8|
|[dataValidationRule](reference/excel/datavalidationrule.md)|_Relationship_ > wholeNumber|WholeNumber data validation criteria.|1.8|
|[dateTimeDataValidation](reference/excel/datetimedatavalidation.md)|_Property_ > formula1|Gets or sets the Formula1, i.e. minimum value or value depending of the operator.|1.8|
|[dateTimeDataValidation](reference/excel/datetimedatavalidation.md)|_Property_ > formula2|Gets or sets the Formula2, i.e. maximum value or value depending of the operator.|1.8|
|[dateTimeDataValidation](reference/excel/datetimedatavalidation.md)|_Relationship_ > operator|The operator to use for validating the data.|1.8|
|[documentProperties](reference/excel/documentproperties.md)|_Property_ > author|Gets or sets the author of the workbook.|1.7|
|[documentProperties](reference/excel/documentproperties.md)|_Property_ > category|Gets or sets the category of the workbook.|1.7|
|[documentProperties](reference/excel/documentproperties.md)|_Property_ > comments|Gets or sets the comments of the workbook.|1.7|
|[documentProperties](reference/excel/documentproperties.md)|_Property_ > company|Gets or sets the company of the workbook.|1.7|
|[documentProperties](reference/excel/documentproperties.md)|_Property_ > keywords|Gets or sets the keywords of the workbook.|1.7|
|[documentProperties](reference/excel/documentproperties.md)|_Property_ > lastAuthor|Gets the last author of the workbook. Read only. Read-only.|1.7|
|[documentProperties](reference/excel/documentproperties.md)|_Property_ > manager|Gets or sets the manager of the workbook.|1.7|
|[documentProperties](reference/excel/documentproperties.md)|_Property_ > revisionNumber|Gets the revision number of the workbook. Read only.|1.7|
|[documentProperties](reference/excel/documentproperties.md)|_Property_ > subject|Gets or sets the subject of the workbook.|1.7|
|[documentProperties](reference/excel/documentproperties.md)|_Property_ > title|Gets or sets the title of the workbook.|1.7|
|[documentProperties](reference/excel/documentproperties.md)|_Relationship_ > creationDate|Gets the creation date of the workbook. Read only. Read-only.|1.7|
|[documentProperties](reference/excel/documentproperties.md)|_Relationship_ > custom|Gets the collection of custom properties of the workbook. Read only. Read-only.|1.7|
|[enableEventsPostProcessAction](reference/excel/enableeventspostprocessaction.md)|_Property_ > isEnableEvents{|Transmits additional data to client side, e.g., worksheetId for TableSelectionChangedEvent.|1.8|
|[enableEventsPostProcessAction](reference/excel/enableeventspostprocessaction.md)|_Relationship_ > actionType|Transmits additional data to client side, e.g., worksheetId for TableSelectionChangedEvent.|1.8|
|[enableEventsPostProcessAction](reference/excel/enableeventspostprocessaction.md)|_Relationship_ > controlId|Transmits additional data to client side, e.g., worksheetId for TableSelectionChangedEvent.|1.8|
|[filterCriteria](reference/excel/filtercriteria.md)|_Property_ > subField|The property used by the filter to do rich filter on richvalues.|beta|
|[filterPivotHierarchy](reference/excel/filterpivothierarchy.md)|_Property_ > enableMultipleFilterItems|Determines whether to allow multiple filter items.|1.8|
|[filterPivotHierarchy](reference/excel/filterpivothierarchy.md)|_Property_ > id|Id of the FilterPivotHierarchy. Read-only.|1.8|
|[filterPivotHierarchy](reference/excel/filterpivothierarchy.md)|_Property_ > name|Name of the FilterPivotHierarchy.|1.8|
|[filterPivotHierarchy](reference/excel/filterpivothierarchy.md)|_Property_ > position|Position of the FilterPivotHierarchy.|1.8|
|[filterPivotHierarchy](reference/excel/filterpivothierarchy.md)|_Relationship_ > fields|Returns the PivotFields associated with the FilterPivotHierarchy. Read-only.|1.8|
|[filterPivotHierarchy](reference/excel/filterpivothierarchy.md)|_Method_ > [setToDefault()]((reference/excel/filterpivothierarchy.md#settodefault)|Reset the FilterPivotHierarchy back to its default values.|1.8|
|[filterPivotHierarchyCollection](reference/excel/filterpivothierarchycollection.md)|_Property_ > items|A collection of filterPivotHierarchy objects. Read-only.|1.8|
|[filterPivotHierarchyCollection](reference/excel/filterpivothierarchycollection.md)|_Method_ > [add(pivotHierarchy: PivotHierarchy)]((reference/excel/filterpivothierarchycollection.md#addpivothierarchy-pivothierarchy)|Adds the PivotHierarchy to the current axis. If the hierarchy is present elsewhere on the row, column,|1.8|
|[filterPivotHierarchyCollection](reference/excel/filterpivothierarchycollection.md)|_Method_ > [getCount()]((reference/excel/filterpivothierarchycollection.md#getcount)|Gets the number of pivot hierarchies in the collection.|1.8|
|[filterPivotHierarchyCollection](reference/excel/filterpivothierarchycollection.md)|_Method_ > [getItem(name: string)]((reference/excel/filterpivothierarchycollection.md#getitemname-string)|Gets a FilterPivotHierarchy by its name or id.|1.8|
|[filterPivotHierarchyCollection](reference/excel/filterpivothierarchycollection.md)|_Method_ > [getItemOrNullObject(name: string)]((reference/excel/filterpivothierarchycollection.md#getitemornullobjectname-string)|Gets a FilterPivotHierarchy by name. If the FilterPivotHierarchy does not exist, will return a null object.|1.8|
|[filterPivotHierarchyCollection](reference/excel/filterpivothierarchycollection.md)|_Method_ > [remove(filterPivotHierarchy: FilterPivotHierarchy)]((reference/excel/filterpivothierarchycollection.md#removefilterpivothierarchy-filterpivothierarchy)|Removes the PivotHierarchy from the current axis.|1.8|
|[geometricShape](reference/excel/geometricshape.md)|_Property_ > id|Represents the shape identifier. Read-only.|beta|
|[geometricShape](reference/excel/geometricshape.md)|_Relationship_ > shape|Returns the shape object for the geometric shape. Read-only.|beta|
|[headerFooter](reference/excel/headerfooter.md)|_Property_ > centerFooter|Gets or sets the center footer of the worksheet.|beta|
|[headerFooter](reference/excel/headerfooter.md)|_Property_ > centerHeader|Gets or sets the center header of the worksheet.|beta|
|[headerFooter](reference/excel/headerfooter.md)|_Property_ > leftFooter|Gets or sets the left footer of the worksheet.|beta|
|[headerFooter](reference/excel/headerfooter.md)|_Property_ > leftHeader|Gets or sets the left header of the worksheet.|beta|
|[headerFooter](reference/excel/headerfooter.md)|_Property_ > rightFooter|Gets or sets the right footer of the worksheet.|beta|
|[headerFooter](reference/excel/headerfooter.md)|_Property_ > rightHeader|Gets or sets the right header of the worksheet.|beta|
|[headerFooterGroup](reference/excel/headerfootergroup.md)|_Property_ > useSheetMargins|Gets or sets a flag indicating if headersfooters are aligned with the page margins set in the page layout options for the worksheet.|beta|
|[headerFooterGroup](reference/excel/headerfootergroup.md)|_Property_ > useSheetScale|Gets or sets a flag indicating if headersfooters should be scaled by the page percentage scale set in the page layout options for the worksheet.|beta|
|[headerFooterGroup](reference/excel/headerfootergroup.md)|_Relationship_ > defaultForAllPages|The general headerfooter, used for all pages unless evenodd or first page is specified. Read-only.|beta|
|[headerFooterGroup](reference/excel/headerfootergroup.md)|_Relationship_ > evenPages|The headerfooter to use for even pages, odd headerfooter needs to be specified for odd pages. Read-only.|beta|
|[headerFooterGroup](reference/excel/headerfootergroup.md)|_Relationship_ > firstPage|The first page headerfooter, for all other pages general or evenodd is used. Read-only.|beta|
|[headerFooterGroup](reference/excel/headerfootergroup.md)|_Relationship_ > oddPages|The headerfooter to use for odd pages, even headerfooter needs to be specified for even pages. Read-only.|beta|
|[headerFooterGroup](reference/excel/headerfootergroup.md)|_Relationship_ > state|Gets or sets the state of which headersfooters are set.|beta|
|[image](reference/excel/image.md)|_Property_ > id|Represents the shape identifier for the image object. Read-only.|beta|
|[image](reference/excel/image.md)|_Relationship_ > shape|Returns the shape object for the image. Read-only.|beta|
|[internalTestEventArgs](reference/excel/internaltesteventargs.md)|_Property_ > prop1|Gets a style by name.|1.7|
|[internalTestEventArgs](reference/excel/internaltesteventargs.md)|_Relationship_ > worksheet|Gets a style by name.|1.7|
|[iterativeCalculation](reference/excel/iterativecalculation.md)|_Property_ > enabled|True if Excel will use iteration to resolve circular references.|beta|
|[iterativeCalculation](reference/excel/iterativecalculation.md)|_Property_ > maxChange|Returns or sets the maximum amount of change between each iteration as Excel resolves circular references.|beta|
|[iterativeCalculation](reference/excel/iterativecalculation.md)|_Property_ > maxIteration|Returns or sets the maximum number of iterations that Excel can use to resolve a circular reference.|beta|
|[listDataValidation](reference/excel/listdatavalidation.md)|_Property_ > inCellDropDown|Displays the list in cell drop down or not, it defaults to true.|1.8|
|[listDataValidation](reference/excel/listdatavalidation.md)|_Property_ > source|Source of the list for data validation|1.8|
|[namedItem](reference/excel/nameditem.md)|_Property_ > formula|Gets or sets the formula of the named item.  Formula always starts with a '=' sign.|1.7|
|[namedItem](reference/excel/nameditem.md)|_Relationship_ > arrayValues|Returns an object containing values and types of the named item. Read-only.|1.7|
|[namedItemArrayValues](reference/excel/nameditemarrayvalues.md)|_Property_ > types|Represents the types for each item in the named item array Read-only. Possible values are: Unknown, Empty, String, Integer, Double, Boolean, Error.|1.7|
|[namedItemArrayValues](reference/excel/nameditemarrayvalues.md)|_Property_ > values|Represents the values of each item in the named item array. Read-only.|1.7|
|[openWorkbookPostProcessAction](reference/excel/openworkbookpostprocessaction.md)|_Property_ > fakeFileId|Transmits additional data to client side, e.g., worksheetId for TableSelectionChangedEvent.|1.8|
|[openWorkbookPostProcessAction](reference/excel/openworkbookpostprocessaction.md)|_Relationship_ > actionType|Transmits additional data to client side, e.g., worksheetId for TableSelectionChangedEvent.|1.8|
|[pageBreak](reference/excel/pagebreak.md)|_Property_ > columnIndex|Represents the column index for the page break Read-only.|beta|
|[pageBreak](reference/excel/pagebreak.md)|_Property_ > rowIndex|Represents the row index for the page break Read-only.|beta|
|[pageBreak](reference/excel/pagebreak.md)|_Method_ > [delete()]((reference/excel/pagebreak.md#delete)|Deletes a page break object.|beta|
|[pageBreak](reference/excel/pagebreak.md)|_Method_ > [getStartCell()]((reference/excel/pagebreak.md#getstartcell)|Gets the first cell after the page break.|beta|
|[pageBreakCollection](reference/excel/pagebreakcollection.md)|_Property_ > items|A collection of pageBreak objects. Read-only.|beta|
|[pageBreakCollection](reference/excel/pagebreakcollection.md)|_Method_ > [add(pageBreakRange: Range or string)]((reference/excel/pagebreakcollection.md#addpagebreakrange-range-or-string)|Adds a page break before the top-left cell of the range specified.|beta|
|[pageBreakCollection](reference/excel/pagebreakcollection.md)|_Method_ > [getCount()]((reference/excel/pagebreakcollection.md#getcount)|Gets the number of page breaks in the collection.|beta|
|[pageBreakCollection](reference/excel/pagebreakcollection.md)|_Method_ > [getItem(index: number)]((reference/excel/pagebreakcollection.md#getitemindex-number)|Gets a page break object via the index.|beta|
|[pageBreakCollection](reference/excel/pagebreakcollection.md)|_Method_ > [removePageBreaks()]((reference/excel/pagebreakcollection.md#removepagebreaks)|Resets all manual page breaks in the collection.|beta|
|[pageLayout](reference/excel/pagelayout.md)|_Property_ > blackAndWhite|Gets or sets the worksheet's black and white print option.|beta|
|[pageLayout](reference/excel/pagelayout.md)|_Property_ > bottomMargin|Gets or sets the worksheet's bottom page margin to use for printing in points.|beta|
|[pageLayout](reference/excel/pagelayout.md)|_Property_ > centerHorizontally|Gets or sets the worksheet's center horizontally flag. This flag determines whether the worksheet will be centered horizontally when it's printed.|beta|
|[pageLayout](reference/excel/pagelayout.md)|_Property_ > centerVertically|Gets or sets the worksheet's center vertically flag. This flag determines whether the worksheet will be centered vertically when it's printed.|beta|
|[pageLayout](reference/excel/pagelayout.md)|_Property_ > draftMode|Gets or sets the worksheet's draft mode option. If true the sheet will be printed without graphics.|beta|
|[pageLayout](reference/excel/pagelayout.md)|_Property_ > firstPageNumber|Gets or sets the worksheet's first page number to print. Null value represents "auto" page numbering.|beta|
|[pageLayout](reference/excel/pagelayout.md)|_Property_ > footerMargin|Gets or sets the worksheet's footer margin, in points, for use when printing.|beta|
|[pageLayout](reference/excel/pagelayout.md)|_Property_ > headerMargin|Gets or sets the worksheet's header margin, in points, for use when printing.|beta|
|[pageLayout](reference/excel/pagelayout.md)|_Property_ > leftMargin|Gets or sets the worksheet's left margin, in points, for use when printing.|beta|
|[pageLayout](reference/excel/pagelayout.md)|_Property_ > orientation|Gets or sets the worksheet's orientation of the page. Possible values are: Portrait, Landscape.|beta|
|[pageLayout](reference/excel/pagelayout.md)|_Property_ > paperSize|Gets or sets the worksheet's paper size of the page. Possible values are: Letter, LetterSmall, Tabloid, Ledger, Legal, Statement, Executive, A3, A4, A4Small, A5, B4, B5, Folio, Quatro, Paper10x14, Paper11x17, Note, Envelope9, Envelope10, Envelope11, Envelope12, Envelope14, Csheet, Dsheet, Esheet, EnvelopeDL, EnvelopeC5, EnvelopeC3, EnvelopeC4, EnvelopeC6, EnvelopeC65, EnvelopeB4, EnvelopeB5, EnvelopeB6, EnvelopeItaly, EnvelopeMonarch, EnvelopePersonal, FanfoldUS, FanfoldStdGerman, FanfoldLegalGerman.|beta|
|[pageLayout](reference/excel/pagelayout.md)|_Property_ > printGridlines|Gets or sets the worksheet's print gridlines flag. This flag determines whether gridlines will be printed or not.|beta|
|[pageLayout](reference/excel/pagelayout.md)|_Property_ > printHeadings|Gets or sets the worksheet's print headings flag. This flag determines whether headings will be printed or not.|beta|
|[pageLayout](reference/excel/pagelayout.md)|_Property_ > rightMargin|Gets or sets the worksheet's right margin, in points, for use when printing.|beta|
|[pageLayout](reference/excel/pagelayout.md)|_Property_ > topMargin|Gets or sets the worksheet's top margin, in points, for use when printing.|beta|
|[pageLayout](reference/excel/pagelayout.md)|_Relationship_ > headersFooters|Header and footer configuration for the worksheet. Read-only.|beta|
|[pageLayout](reference/excel/pagelayout.md)|_Relationship_ > printComments|Gets or sets whether the worksheet's comments should be displayed when printing.|beta|
|[pageLayout](reference/excel/pagelayout.md)|_Relationship_ > printErrors|Gets or sets the worksheet's print errors option.|beta|
|[pageLayout](reference/excel/pagelayout.md)|_Relationship_ > printOrder|Gets or sets the worksheet's page print order option. This specifies the order to use for processing the page number printed.|beta|
|[pageLayout](reference/excel/pagelayout.md)|_Relationship_ > zoom|Gets or sets the worksheet's print zoom options.|beta|
|[pageLayout](reference/excel/pagelayout.md)|_Method_ > [getPrintArea()]((reference/excel/pagelayout.md#getprintarea)|Gets the RangeAreas object, comprising one or more rectangular ranges, that represents the print area for the worksheet. If there is no print area, an ItemNotFound error will be thrown.|beta|
|[pageLayout](reference/excel/pagelayout.md)|_Method_ > [getPrintAreaOrNullObject()]((reference/excel/pagelayout.md#getprintareaornullobject)|Gets the RangeAreas object, comprising one or more rectangular ranges, that represents the print area for the worksheet. If there is no print area, a null object will be returned.|beta|
|[pageLayout](reference/excel/pagelayout.md)|_Method_ > [getPrintTitleColumns()]((reference/excel/pagelayout.md#getprinttitlecolumns)|Gets the range object representing the title columns.|beta|
|[pageLayout](reference/excel/pagelayout.md)|_Method_ > [getPrintTitleColumnsOrNullObject()]((reference/excel/pagelayout.md#getprinttitlecolumnsornullobject)|Gets the range object representing the title columns. If not set, this will return a null object.|beta|
|[pageLayout](reference/excel/pagelayout.md)|_Method_ > [getPrintTitleRows()]((reference/excel/pagelayout.md#getprinttitlerows)|Gets the range object representing the title rows.|beta|
|[pageLayout](reference/excel/pagelayout.md)|_Method_ > [getPrintTitleRowsOrNullObject()]((reference/excel/pagelayout.md#getprinttitlerowsornullobject)|Gets the range object representing the title rows. If not set, this will return a null object.|beta|
|[pageLayout](reference/excel/pagelayout.md)|_Method_ > [setPrintArea(printArea: Range or RangeAreas or string)]((reference/excel/pagelayout.md#setprintareaprintarea-range-or-rangeareas-or-string)|Sets the worksheet's print area.|beta|
|[pageLayout](reference/excel/pagelayout.md)|_Method_ > [setPrintMargins(unit: PrintMarginUnit, marginOptions: PageLayoutMarginOptions)]((reference/excel/pagelayout.md#setprintmarginsunit-printmarginunit-marginoptions-pagelayoutmarginoptions)|Sets the worksheet's page margins with units.|beta|
|[pageLayout](reference/excel/pagelayout.md)|_Method_ > [setPrintTitleColumns(printTitleColumns: Range or string)]((reference/excel/pagelayout.md#setprinttitlecolumnsprinttitlecolumns-range-or-string)|Sets the columns that contain the cells to be repeated at the left of each page of the worksheet for printing.|beta|
|[pageLayout](reference/excel/pagelayout.md)|_Method_ > [setPrintTitleRows(printTitleRows: Range or string)]((reference/excel/pagelayout.md#setprinttitlerowsprinttitlerows-range-or-string)|Sets the rows that contain the cells to be repeated at the top of each page of the worksheet for printing.|beta|
|[pageLayoutMarginOptions](reference/excel/pagelayoutmarginoptions.md)|_Property_ > bottom|Represents the page layout bottom margin in the unit specified to use for printing.|beta|
|[pageLayoutMarginOptions](reference/excel/pagelayoutmarginoptions.md)|_Property_ > footer|Represents the page layout footer margin in the unit specified to use for printing.|beta|
|[pageLayoutMarginOptions](reference/excel/pagelayoutmarginoptions.md)|_Property_ > header|Represents the page layout header margin in the unit specified to use for printing.|beta|
|[pageLayoutMarginOptions](reference/excel/pagelayoutmarginoptions.md)|_Property_ > left|Represents the page layout left margin in the unit specified to use for printing.|beta|
|[pageLayoutMarginOptions](reference/excel/pagelayoutmarginoptions.md)|_Property_ > right|Represents the page layout right margin in the unit specified to use for printing.|beta|
|[pageLayoutMarginOptions](reference/excel/pagelayoutmarginoptions.md)|_Property_ > top|Represents the page layout top margin in the unit specified to use for printing.|beta|
|[pageLayoutZoomOptions](reference/excel/pagelayoutzoomoptions.md)|_Property_ > horizontalFitToPages|Number of pages to fit horizontally. This value can be null if percentage scale is used.|beta|
|[pageLayoutZoomOptions](reference/excel/pagelayoutzoomoptions.md)|_Property_ > scale|Print page scale value can be between 10 and 400. This value can be null if fit to page tall or wide is specified.|beta|
|[pageLayoutZoomOptions](reference/excel/pagelayoutzoomoptions.md)|_Property_ > verticalFitToPages|Number of pages to fit vertically. This value can be null if percentage scale is used.|beta|
|[pivotField](reference/excel/pivotfield.md)|_Property_ > id|Id of the PivotField. Read-only.|1.8|
|[pivotField](reference/excel/pivotfield.md)|_Property_ > name|Name of the PivotField.|1.8|
|[pivotField](reference/excel/pivotfield.md)|_Property_ > showAllItems|Determines whether to show all items of the PivotField.|1.8|
|[pivotField](reference/excel/pivotfield.md)|_Relationship_ > items|Returns the PivotFields associated with the PivotField. Read-only.|1.8|
|[pivotField](reference/excel/pivotfield.md)|_Relationship_ > subtotals|Subtotals of the PivotField.|1.8|
|[pivotField](reference/excel/pivotfield.md)|_Method_ > [sortByLabels(sortby: SortBy)]((reference/excel/pivotfield.md#sortbylabelssortby-sortby)|Sorts the PivotField. If a DataPivotHierarchy is specified, then sort will be applied based on it, if not sort will be based on the PivotField itself.|1.8|
|[pivotField](reference/excel/pivotfield.md)|_Method_ > [sortByValues(sortby: SortBy, valuesHierarchy: DataPivotHierarchy, pivotItemScope: ()[])]((reference/excel/pivotfield.md#sortbyvaluessortby-sortby-valueshierarchy-datapivothierarchy-pivotitemscope-)|Sorts the PivotField by specified values in a given scope. The scope defines which specific values will be used to sort when|beta|
|[pivotFieldCollection](reference/excel/pivotfieldcollection.md)|_Property_ > items|A collection of pivotField objects. Read-only.|1.8|
|[pivotFieldCollection](reference/excel/pivotfieldcollection.md)|_Method_ > [getCount()]((reference/excel/pivotfieldcollection.md#getcount)|Gets the number of pivot hierarchies in the collection.|1.8|
|[pivotFieldCollection](reference/excel/pivotfieldcollection.md)|_Method_ > [getItem(name: string)]((reference/excel/pivotfieldcollection.md#getitemname-string)|Gets a PivotHierarchy by its name or id.|1.8|
|[pivotFieldCollection](reference/excel/pivotfieldcollection.md)|_Method_ > [getItemOrNullObject(name: string)]((reference/excel/pivotfieldcollection.md#getitemornullobjectname-string)|Gets a PivotHierarchy by name. If the PivotHierarchy does not exist, will return a null object.|1.8|
|[pivotHierarchy](reference/excel/pivothierarchy.md)|_Property_ > id|Id of the PivotHierarchy. Read-only.|1.8|
|[pivotHierarchy](reference/excel/pivothierarchy.md)|_Property_ > name|Name of the PivotHierarchy.|1.8|
|[pivotHierarchy](reference/excel/pivothierarchy.md)|_Relationship_ > fields|Returns the PivotFields associated with the PivotHierarchy. Read-only.|1.8|
|[pivotHierarchyCollection](reference/excel/pivothierarchycollection.md)|_Property_ > items|A collection of pivotHierarchy objects. Read-only.|1.8|
|[pivotHierarchyCollection](reference/excel/pivothierarchycollection.md)|_Method_ > [getCount()]((reference/excel/pivothierarchycollection.md#getcount)|Gets the number of pivot hierarchies in the collection.|1.8|
|[pivotHierarchyCollection](reference/excel/pivothierarchycollection.md)|_Method_ > [getItem(name: string)]((reference/excel/pivothierarchycollection.md#getitemname-string)|Gets a PivotHierarchy by its name or id.|1.8|
|[pivotHierarchyCollection](reference/excel/pivothierarchycollection.md)|_Method_ > [getItemOrNullObject(name: string)]((reference/excel/pivothierarchycollection.md#getitemornullobjectname-string)|Gets a PivotHierarchy by name. If the PivotHierarchy does not exist, will return a null object.|1.8|
|[pivotItem](reference/excel/pivotitem.md)|_Property_ > id|Id of the PivotItem. Read-only.|1.8|
|[pivotItem](reference/excel/pivotitem.md)|_Property_ > isExpanded|Determines whether the item is expanded to show child items or if it's collapsed and child items are hidden.|1.8|
|[pivotItem](reference/excel/pivotitem.md)|_Property_ > name|Name of the PivotItem.|1.8|
|[pivotItem](reference/excel/pivotitem.md)|_Property_ > visible|Determines whether the PivotItem is visible or not.|1.8|
|[pivotItemCollection](reference/excel/pivotitemcollection.md)|_Property_ > items|A collection of pivotItem objects. Read-only.|1.8|
|[pivotItemCollection](reference/excel/pivotitemcollection.md)|_Method_ > [getCount()]((reference/excel/pivotitemcollection.md#getcount)|Gets the number of pivot hierarchies in the collection.|1.8|
|[pivotItemCollection](reference/excel/pivotitemcollection.md)|_Method_ > [getItem(name: string)]((reference/excel/pivotitemcollection.md#getitemname-string)|Gets a PivotHierarchy by its name or id.|1.8|
|[pivotItemCollection](reference/excel/pivotitemcollection.md)|_Method_ > [getItemOrNullObject(name: string)]((reference/excel/pivotitemcollection.md#getitemornullobjectname-string)|Gets a PivotHierarchy by name. If the PivotHierarchy does not exist, will return a null object.|1.8|
|[pivotLayout](reference/excel/pivotlayout.md)|_Property_ > enableFieldList|True if the field list should be shown or hidden from the UI.|beta|
|[pivotLayout](reference/excel/pivotlayout.md)|_Property_ > showColumnGrandTotals|True if the PivotTable report shows grand totals for columns.|1.8|
|[pivotLayout](reference/excel/pivotlayout.md)|_Property_ > showRowGrandTotals|True if the PivotTable report shows grand totals for rows.|1.8|
|[pivotLayout](reference/excel/pivotlayout.md)|_Property_ > subtotalLocation|This property indicates the SubtotalLocationType of all fields on the PivotTable. If fields have different states, this will be null. Possible values are: AtTop, AtBottom.|1.8|
|[pivotLayout](reference/excel/pivotlayout.md)|_Relationship_ > layoutType|This property indicates the PivotLayoutType of all fields on the PivotTable. If fields have different states, this will be null.|1.8|
|[pivotLayout](reference/excel/pivotlayout.md)|_Method_ > [getColumnLabelRange()]((reference/excel/pivotlayout.md#getcolumnlabelrange)|Returns the range where the PivotTable's column labels reside.|1.8|
|[pivotLayout](reference/excel/pivotlayout.md)|_Method_ > [getDataBodyRange()]((reference/excel/pivotlayout.md#getdatabodyrange)|Returns the range where the PivotTable's data values reside.|1.8|
|[pivotLayout](reference/excel/pivotlayout.md)|_Method_ > [getFilterAxisRange()]((reference/excel/pivotlayout.md#getfilteraxisrange)|Returns the range of the PivotTable's filter area.|1.8|
|[pivotLayout](reference/excel/pivotlayout.md)|_Method_ > [getRange()]((reference/excel/pivotlayout.md#getrange)|Returns the range the PivotTable exists on, excluding the filter area.|1.8|
|[pivotLayout](reference/excel/pivotlayout.md)|_Method_ > [getRowLabelRange()]((reference/excel/pivotlayout.md#getrowlabelrange)|Returns the range where the PivotTable's row labels reside.|1.8|
|[pivotTable](reference/excel/pivottable.md)|_Property_ > useCustomSortLists|True if the PivotTable should use custom lists when sorting.|beta|
|[pivotTable](reference/excel/pivottable.md)|_Relationship_ > columnHierarchies|The Column Pivot Hierarchies of the PivotTable. Read-only.|1.8|
|[pivotTable](reference/excel/pivottable.md)|_Relationship_ > dataHierarchies|The Data Pivot Hierarchies of the PivotTable. Read-only.|1.8|
|[pivotTable](reference/excel/pivottable.md)|_Relationship_ > filterHierarchies|The Filter Pivot Hierarchies of the PivotTable. Read-only.|1.8|
|[pivotTable](reference/excel/pivottable.md)|_Relationship_ > hierarchies|The Pivot Hierarchies of the PivotTable. Read-only.|1.8|
|[pivotTable](reference/excel/pivottable.md)|_Relationship_ > layout|The PivotLayout describing the layout and visual structure of the PivotTable. Read-only.|1.8|
|[pivotTable](reference/excel/pivottable.md)|_Relationship_ > rowHierarchies|The Row Pivot Hierarchies of the PivotTable. Read-only.|1.8|
|[pivotTable](reference/excel/pivottable.md)|_Method_ > [delete()]((reference/excel/pivottable.md#delete)|Deletes the PivotTable.|1.8|
|[pivotTableCollection](reference/excel/pivottablecollection.md)|_Method_ > [add(name: string, source: object, destination: object)]((reference/excel/pivottablecollection.md#addname-string-source-object-destination-object)|Add a Pivottable based on the specified source data and insert it at the top left cell of the destination range.|1.8|
|[range](reference/excel/range.md)|_Property_ > isEntireColumn|Represents if the current range is an entire column. Read-only.|1.7|
|[range](reference/excel/range.md)|_Property_ > isEntireRow|Represents if the current range is an entire row. Read-only.|1.7|
|[range](reference/excel/range.md)|_Property_ > numberFormatLocal|Represents Excel's number format code for the given range as a string in the language of the user.|1.7|
|[range](reference/excel/range.md)|_Property_ > style|Represents the style of the current range.|1.7|
|[range](reference/excel/range.md)|_Relationship_ > dataValidation|Returns a data validation object. Read-only.|1.8|
|[range](reference/excel/range.md)|_Relationship_ > hyperlink|Represents the hyperlink for the current range.|1.7|
|[range](reference/excel/range.md)|_Relationship_ > linkedDataTypeState|Represents the data type state of each cell. Read-only.|beta|
|[range](reference/excel/range.md)|_Method_ > [convertDataTypeToText()]((reference/excel/range.md#convertdatatypetotext)|Converts the range cells with datatypes into text.|beta|
|[range](reference/excel/range.md)|_Method_ > [convertToLinkedDataType(serviceID: number, languageCulture: string)]((reference/excel/range.md#converttolinkeddatatypeserviceid-number-languageculture-string)|Converts the range cells into linked datatype in the worksheet.|beta|
|[range](reference/excel/range.md)|_Method_ > [copyFrom(sourceRange: Range or RangeAreas or string, copyType: RangeCopyType, skipBlanks: bool, transpose: bool)]((reference/excel/range.md#copyfromsourcerange-range-or-rangeareas-or-string-copytype-rangecopytype-skipblanks-bool-transpose-bool)|Copies cell data or formatting from the source range or RangeAreas to the current range.|beta|
|[range](reference/excel/range.md)|_Method_ > [find(text: string, criteria: SearchCriteria)]((reference/excel/range.md#findtext-string-criteria-searchcriteria)|Finds the given string based on the criteria specified.|beta|
|[range](reference/excel/range.md)|_Method_ > [findOrNullObject(text: string, criteria: SearchCriteria)]((reference/excel/range.md#findornullobjecttext-string-criteria-searchcriteria)|Finds the given string based on the criteria specified.|beta|
|[range](reference/excel/range.md)|_Method_ > [getAbsoluteResizedRange(numRows: number, numColumns: number)]((reference/excel/range.md#getabsoluteresizedrangenumrows-number-numcolumns-number)|Gets a Range object with the same top-left cell as the current Range object, but with the specified numbers of rows and columns.|1.7|
|[range](reference/excel/range.md)|_Method_ > [getImage()]((reference/excel/range.md#getimage)|Renders the range as a base64-encoded png image.|1.7|
|[range](reference/excel/range.md)|_Method_ > [getSpecialCells(cellType: SpecialCellType, cellValueType: SpecialCellValueType)]((reference/excel/range.md#getspecialcellscelltype-specialcelltype-cellvaluetype-specialcellvaluetype)|Gets the RangeAreas object, comprising one or more rectangular ranges, that represents all the cells that match the specified type and value.|beta|
|[range](reference/excel/range.md)|_Method_ > [getSpecialCellsOrNullObject(cellType: SpecialCellType, cellValueType: SpecialCellValueType)]((reference/excel/range.md#getspecialcellsornullobjectcelltype-specialcelltype-cellvaluetype-specialcellvaluetype)|Gets the RangeAreas object, comprising one or more ranges, that represents all the cells that match the specified type and value.|beta|
|[range](reference/excel/range.md)|_Method_ > [getSurroundingRegion()]((reference/excel/range.md#getsurroundingregion)|Returns a Range object that represents the surrounding region for the top-left cell in this range. A surrounding region is a range bounded by any combination of blank rows and blank columns relative to this range.|1.7|
|[range](reference/excel/range.md)|_Method_ > [getTables(fullyContained: bool)]((reference/excel/range.md#gettablesfullycontained-bool)|Gets a scoped collection of tables that overlap with the range.|beta|
|[range](reference/excel/range.md)|_Method_ > [removeDuplicates(columns: int[], includesHeader: bool)]((reference/excel/range.md#removeduplicatescolumns-int-includesheader-bool)|Removes duplicate values from the range specified by the columns.|beta|
|[range](reference/excel/range.md)|_Method_ > [replaceAll(text: string, replacement: string, criteria: ReplaceCriteria)]((reference/excel/range.md#replacealltext-string-replacement-string-criteria-replacecriteria)|Finds and replaces the given string based on the criteria specified within the current range.|beta|
|[range](reference/excel/range.md)|_Method_ > [setDirty()]((reference/excel/range.md#setdirty)|Set a range to be recalculated when the next recalculation occurs.|beta|
|[range](reference/excel/range.md)|_Method_ > [showCard()]((reference/excel/range.md#showcard)|Displays the card for an active cell if it has rich value content.|1.7|
|[rangeAreas](reference/excel/rangeareas.md)|_Property_ > address|Returns the RageAreas reference in A1-style. Address value will contain the worksheet name for each rectangular block of cells (e.g. "Sheet1!A1:B4, Sheet1!D1:D4"). Read-only.|beta|
|[rangeAreas](reference/excel/rangeareas.md)|_Property_ > addressLocal|Returns the RageAreas reference in the user locale. Read-only.|beta|
|[rangeAreas](reference/excel/rangeareas.md)|_Property_ > areaCount|Returns the number of rectangular ranges that comprise this RangeAreas object. Read-only.|beta|
|[rangeAreas](reference/excel/rangeareas.md)|_Property_ > cellCount|Returns the number of cells in the RangeAreas object, summing up the cell counts of all of the individual rectangular ranges. Returns -1 if the cell count exceeds 2^31-1 (2,147,483,647). Read-only.|beta|
|[rangeAreas](reference/excel/rangeareas.md)|_Property_ > isEntireColumn|Indicates whether all the ranges on this RangeAreas object represent entire columns (e.g., "A:C, Q:Z"). Read-only.|beta|
|[rangeAreas](reference/excel/rangeareas.md)|_Property_ > isEntireRow|Indicates whether all the ranges on this RangeAreas object represent entire rows (e.g., "1:3, 5:7"). Read-only.|beta|
|[rangeAreas](reference/excel/rangeareas.md)|_Property_ > style|Represents the style for all ranges in this RangeAreas object.|beta|
|[rangeAreas](reference/excel/rangeareas.md)|_Relationship_ > areas|Returns a collection of rectangular ranges that comprise this RangeAreas object. Read-only.|beta|
|[rangeAreas](reference/excel/rangeareas.md)|_Relationship_ > conditionalFormats|Returns a collection of ConditionalFormats that intersect with any cells in this RangeAreas object. Read-only.|beta|
|[rangeAreas](reference/excel/rangeareas.md)|_Relationship_ > dataValidation|Returns a dataValidation object for all ranges in the RangeAreas. Read-only.|beta|
|[rangeAreas](reference/excel/rangeareas.md)|_Relationship_ > format|Returns a rangeFormat object, encapsulating the the font, fill, borders, alignment, and other properties for all ranges in the RangeAreas object. Read-only.|beta|
|[rangeAreas](reference/excel/rangeareas.md)|_Relationship_ > worksheet|Returns the worksheet for the current RangeAreas. Read-only.|beta|
|[rangeAreas](reference/excel/rangeareas.md)|_Method_ > [calculate()]((reference/excel/rangeareas.md#calculate)|Calculates all cells in the RangeAreas.|beta|
|[rangeAreas](reference/excel/rangeareas.md)|_Method_ > [clear(applyTo: string)]((reference/excel/rangeareas.md#clearapplyto-string)|Clears values, format, fill, border, etc on each of the areas that comprise this RangeAreas object.|beta|
|[rangeAreas](reference/excel/rangeareas.md)|_Method_ > [convertDataTypeToText()]((reference/excel/rangeareas.md#convertdatatypetotext)|Converts all cells in the RangeAreas with datatypes into text.|beta|
|[rangeAreas](reference/excel/rangeareas.md)|_Method_ > [convertToLinkedDataType(serviceID: number, languageCulture: string)]((reference/excel/rangeareas.md#converttolinkeddatatypeserviceid-number-languageculture-string)|Converts all cells in the RangeAreas into linked datatype.|beta|
|[rangeAreas](reference/excel/rangeareas.md)|_Method_ > [copyFrom(sourceRange: Range or RangeAreas or string, copyType: RangeCopyType, skipBlanks: bool, transpose: bool)]((reference/excel/rangeareas.md#copyfromsourcerange-range-or-rangeareas-or-string-copytype-rangecopytype-skipblanks-bool-transpose-bool)|Copies cell data or formatting from the source range or RangeAreas to the current RangeAreas.|beta|
|[rangeAreas](reference/excel/rangeareas.md)|_Method_ > [getEntireColumn()]((reference/excel/rangeareas.md#getentirecolumn)|Returns a RangeAreas object that represents the entire columns of the RangeAreas (for example, if the current RangeAreas represents cells "B4:E11, H2", it returns a RangeAreas that represents columns "B:E, H:H").|beta|
|[rangeAreas](reference/excel/rangeareas.md)|_Method_ > [getEntireRow()]((reference/excel/rangeareas.md#getentirerow)|Returns a RangeAreas object that represents the entire rows of the RangeAreas (for example, if the current RangeAreas represents cells "B4:E11", it returns a RangeAreas that represents rows "4:11").|beta|
|[rangeAreas](reference/excel/rangeareas.md)|_Method_ > [getIntersection(anotherRange: Range or RangeAreas or string)]((reference/excel/rangeareas.md#getintersectionanotherrange-range-or-rangeareas-or-string)|Returns the RangeAreas object that represents the intersection of the given ranges or RangeAreas. If no intersection is found, an ItemNotFound error will be thrown.|beta|
|[rangeAreas](reference/excel/rangeareas.md)|_Method_ > [getIntersectionOrNullObject(anotherRange: Range or RangeAreas or string)]((reference/excel/rangeareas.md#getintersectionornullobjectanotherrange-range-or-rangeareas-or-string)|Returns the RangeAreas object that represents the intersection of the given ranges or RangeAreas. If no intersection is found, a null object is returned.|beta|
|[rangeAreas](reference/excel/rangeareas.md)|_Method_ > [getOffsetRangeAreas(rowOffset: number, columnOffset: number)]((reference/excel/rangeareas.md#getoffsetrangeareasrowoffset-number-columnoffset-number)|Returns an RangeAreas object that is shifted by the specific row and column offset. The dimension of the returned RangeAreas will match the original object. If the resulting RangeAreas is forced outside the bounds of the worksheet grid, an error will be thrown.|beta|
|[rangeAreas](reference/excel/rangeareas.md)|_Method_ > [getSpecialCells(cellType: SpecialCellType, cellValueType: SpecialCellValueType)]((reference/excel/rangeareas.md#getspecialcellscelltype-specialcelltype-cellvaluetype-specialcellvaluetype)|Returns a RangeAreas object that represents all the cells that match the specified type and value. Throws an error if no special cells are found that match the criteria.|beta|
|[rangeAreas](reference/excel/rangeareas.md)|_Method_ > [getSpecialCellsOrNullObject(cellType: SpecialCellType, cellValueType: SpecialCellValueType)]((reference/excel/rangeareas.md#getspecialcellsornullobjectcelltype-specialcelltype-cellvaluetype-specialcellvaluetype)|Returns a RangeAreas object that represents all the cells that match the specified type and value. Returns a null object if no special cells are found that match the criteria.|beta|
|[rangeAreas](reference/excel/rangeareas.md)|_Method_ > [getTables(fullyContained: bool)]((reference/excel/rangeareas.md#gettablesfullycontained-bool)|Returns a scoped collection of tables that overlap with any range in this RangeAreas object.|beta|
|[rangeAreas](reference/excel/rangeareas.md)|_Method_ > [getUsedRangeAreas(valuesOnly: bool)]((reference/excel/rangeareas.md#getusedrangeareasvaluesonly-bool)|Returns the used RangeAreas that comprises all the used areas of individual rectangular ranges in the RangeAreas object.|beta|
|[rangeAreas](reference/excel/rangeareas.md)|_Method_ > [getUsedRangeAreasOrNullObject(valuesOnly: bool)]((reference/excel/rangeareas.md#getusedrangeareasornullobjectvaluesonly-bool)|Returns the used RangeAreas that comprises all the used areas of individual rectangular ranges in the RangeAreas object.|beta|
|[rangeAreas](reference/excel/rangeareas.md)|_Method_ > [setDirty()]((reference/excel/rangeareas.md#setdirty)|Sets the RangeAreas to be recalculated when the next recalculation occurs.|beta|
|[rangeBorder](reference/excel/rangeborder.md)|_Property_ > tintAndShade|Returns or sets a double that lightens or darkens a color for Range Border, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.|beta|
|[rangeBorderCollection](reference/excel/rangebordercollection.md)|_Property_ > tintAndShade|Returns or sets a double that lightens or darkens a color for Range Borders, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.|beta|
|[rangeCollection](reference/excel/rangecollection.md)|_Property_ > items|A collection of range objects. Read-only.|beta|
|[rangeCollection](reference/excel/rangecollection.md)|_Method_ > [getCount()]((reference/excel/rangecollection.md#getcount)|Returns the number of ranges in the RangeCollection.|beta|
|[rangeCollection](reference/excel/rangecollection.md)|_Method_ > [getItemAt(index: number)]((reference/excel/rangecollection.md#getitematindex-number)|Returns the range object based on its position in the RangeCollection.|beta|
|[rangeFill](reference/excel/rangefill.md)|_Property_ > patternColor|Sets HTML color code representing the color of the Range pattern, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|beta|
|[rangeFill](reference/excel/rangefill.md)|_Property_ > patternTintAndShade|Returns or sets a double that lightens or darkens a pattern color for Range Fill, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.|beta|
|[rangeFill](reference/excel/rangefill.md)|_Property_ > tintAndShade|Returns or sets a double that lightens or darkens a color for Range Fill, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.|beta|
|[rangeFill](reference/excel/rangefill.md)|_Relationship_ > pattern|Gets or sets the pattern of a Range.|beta|
|[rangeFont](reference/excel/rangefont.md)|_Property_ > strikethrough|Represents the strikethrough status of font. A null value indicates that the entire range doesn't have uniform Strikethrough setting.|beta|
|[rangeFont](reference/excel/rangefont.md)|_Property_ > subscript|Represents the Subscript status of font.|beta|
|[rangeFont](reference/excel/rangefont.md)|_Property_ > superscript|Represents the Superscript status of font.|beta|
|[rangeFont](reference/excel/rangefont.md)|_Property_ > tintAndShade|Returns or sets a double that lightens or darkens a color for Range Font, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.|beta|
|[rangeFormat](reference/excel/rangeformat.md)|_Property_ > autoIndent|Indicates if text is automatically indented when text alignment is set to equal distribution.|beta|
|[rangeFormat](reference/excel/rangeformat.md)|_Property_ > indentLevel|An integer from 0 to 250 that indicates the indent level.|beta|
|[rangeFormat](reference/excel/rangeformat.md)|_Property_ > readingOrder|The reading order for the range. Possible values are: Context, LeftToRight, RightToLeft.|beta|
|[rangeFormat](reference/excel/rangeformat.md)|_Property_ > shrinkToFit|Indicates if text automatically shrinks to fit in the available column width.|beta|
|[rangeFormat](reference/excel/rangeformat.md)|_Property_ > textOrientation|Gets or sets the text orientation of all the cells within the range.|1.7|
|[rangeFormat](reference/excel/rangeformat.md)|_Property_ > useStandardHeight|Determines if the row height of the Range object equals the standard height of the sheet.|1.7|
|[rangeFormat](reference/excel/rangeformat.md)|_Property_ > useStandardWidth|Indicates whether the column width of the Range object equals the standard width of the sheet.|1.7|
|[rangeHyperlink](reference/excel/rangehyperlink.md)|_Property_ > address|Represents the url target for the hyperlink.|1.7|
|[rangeHyperlink](reference/excel/rangehyperlink.md)|_Property_ > documentReference|Represents the document reference target for the hyperlink.|1.7|
|[rangeHyperlink](reference/excel/rangehyperlink.md)|_Property_ > screenTip|Represents the string displayed when hovering over the hyperlink.|1.7|
|[rangeHyperlink](reference/excel/rangehyperlink.md)|_Property_ > textToDisplay|Represents the string that is displayed in the top left most cell in the range.|1.7|
|[rangeViewCollection](reference/excel/rangeviewcollection.md)|_Method_ > [getFirst()]((reference/excel/rangeviewcollection.md#getfirst)|Gets the first RangeView object in the collection.|ApiSetAttribute.Spec|
|[rangeViewCollection](reference/excel/rangeviewcollection.md)|_Method_ > [getLast()]((reference/excel/rangeviewcollection.md#getlast)|Gets the last RangeView object in the collection.|ApiSetAttribute.Spec|
|[registerEventPostProcessAction](reference/excel/registereventpostprocessaction.md)|_Property_ > message|Transmits additional data to client side, e.g., worksheetId for TableSelectionChangedEvent.|1.7|
|[registerEventPostProcessAction](reference/excel/registereventpostprocessaction.md)|_Property_ > targetId|Gets the right border.|1.7|
|[registerEventPostProcessAction](reference/excel/registereventpostprocessaction.md)|_Relationship_ > actionType|Gets the right border.|1.7|
|[registerEventPostProcessAction](reference/excel/registereventpostprocessaction.md)|_Relationship_ > controlId|Gets the right border.|1.7|
|[registerEventPostProcessAction](reference/excel/registereventpostprocessaction.md)|_Relationship_ > messageType|Gets the right border.|1.7|
|[removeDuplicatesResult](reference/excel/removeduplicatesresult.md)|_Property_ > removed|Number of duplicated rows removed by the operation. Read-only.|beta|
|[removeDuplicatesResult](reference/excel/removeduplicatesresult.md)|_Property_ > uniqueRemaining|Number of remaining unique rows present in the resulting range. Read-only.|beta|
|[replaceCriteria](reference/excel/replacecriteria.md)|_Property_ > completeMatch|Specifies whether the match needs to be complete or partial. Default is false (partial).|beta|
|[replaceCriteria](reference/excel/replacecriteria.md)|_Property_ > matchCase|Specifies whether the match is case sensitive. Default is false (insensitive).|beta|
|[rowColumnPivotHierarchy](reference/excel/rowcolumnpivothierarchy.md)|_Property_ > id|Id of the RowColumnPivotHierarchy. Read-only.|1.8|
|[rowColumnPivotHierarchy](reference/excel/rowcolumnpivothierarchy.md)|_Property_ > name|Name of the RowColumnPivotHierarchy.|1.8|
|[rowColumnPivotHierarchy](reference/excel/rowcolumnpivothierarchy.md)|_Property_ > position|Position of the RowColumnPivotHierarchy.|1.8|
|[rowColumnPivotHierarchy](reference/excel/rowcolumnpivothierarchy.md)|_Relationship_ > fields|Returns the PivotFields associated with the RowColumnPivotHierarchy. Read-only.|1.8|
|[rowColumnPivotHierarchy](reference/excel/rowcolumnpivothierarchy.md)|_Method_ > [setToDefault()]((reference/excel/rowcolumnpivothierarchy.md#settodefault)|Reset the RowColumnPivotHierarchy back to its default values.|1.8|
|[rowColumnPivotHierarchyCollection](reference/excel/rowcolumnpivothierarchycollection.md)|_Property_ > items|A collection of rowColumnPivotHierarchy objects. Read-only.|1.8|
|[rowColumnPivotHierarchyCollection](reference/excel/rowcolumnpivothierarchycollection.md)|_Method_ > [add(pivotHierarchy: PivotHierarchy)]((reference/excel/rowcolumnpivothierarchycollection.md#addpivothierarchy-pivothierarchy)|Adds the PivotHierarchy to the current axis. If the hierarchy is present elsewhere on the row, column,|1.8|
|[rowColumnPivotHierarchyCollection](reference/excel/rowcolumnpivothierarchycollection.md)|_Method_ > [getCount()]((reference/excel/rowcolumnpivothierarchycollection.md#getcount)|Gets the number of pivot hierarchies in the collection.|1.8|
|[rowColumnPivotHierarchyCollection](reference/excel/rowcolumnpivothierarchycollection.md)|_Method_ > [getItem(name: string)]((reference/excel/rowcolumnpivothierarchycollection.md#getitemname-string)|Gets a RowColumnPivotHierarchy by its name or id.|1.8|
|[rowColumnPivotHierarchyCollection](reference/excel/rowcolumnpivothierarchycollection.md)|_Method_ > [getItemOrNullObject(name: string)]((reference/excel/rowcolumnpivothierarchycollection.md#getitemornullobjectname-string)|Gets a RowColumnPivotHierarchy by name. If the RowColumnPivotHierarchy does not exist, will return a null object.|1.8|
|[rowColumnPivotHierarchyCollection](reference/excel/rowcolumnpivothierarchycollection.md)|_Method_ > [remove(rowColumnPivotHierarchy: RowColumnPivotHierarchy)]((reference/excel/rowcolumnpivothierarchycollection.md#removerowcolumnpivothierarchy-rowcolumnpivothierarchy)|Removes the PivotHierarchy from the current axis.|1.8|
|[runtime](reference/excel/runtime.md)|_Property_ > enableEvents|Turn onoff JavaScript events in current taskpane or content add-in.|1.8|
|[searchCriteria](reference/excel/searchcriteria.md)|_Property_ > completeMatch|Specifies whether the match needs to be complete or partial. Default is false (partial).|beta|
|[searchCriteria](reference/excel/searchcriteria.md)|_Property_ > matchCase|Specifies whether the match is case sensitive. Default is false (insensitive).|beta|
|[searchCriteria](reference/excel/searchcriteria.md)|_Relationship_ > searchDirection|Specifies the search direction. Default is forward.|beta|
|[shape](reference/excel/shape.md)|_Property_ > altTextDescription|Returns or sets the alternative descriptive text string for a Shape object when the object is saved to a Web page.|beta|
|[shape](reference/excel/shape.md)|_Property_ > altTextTitle|Returns or sets the alternative title text string for a Shape object when the object is saved to a Web page.|beta|
|[shape](reference/excel/shape.md)|_Property_ > height|Represents the height, in points, of the shape.|beta|
|[shape](reference/excel/shape.md)|_Property_ > id|Represents the shape identifier. Read-only.|beta|
|[shape](reference/excel/shape.md)|_Property_ > left|The distance, in points, from the left side of the shape to the left of the worksheet.|beta|
|[shape](reference/excel/shape.md)|_Property_ > lockAspectRatio|Represents if the aspect ratio locked, in boolean, of the shape.|beta|
|[shape](reference/excel/shape.md)|_Property_ > name|Represents the name of the shape. Read-only.|beta|
|[shape](reference/excel/shape.md)|_Property_ > rotation|Represents the rotation, in degrees, of the shape.|beta|
|[shape](reference/excel/shape.md)|_Property_ > top|The distance, in points, from the top edge of the shape to the top of the worksheet.|beta|
|[shape](reference/excel/shape.md)|_Property_ > width|Represents the width, in points, of the shape.|beta|
|[shape](reference/excel/shape.md)|_Property_ > zOrderPosition|Returns the position of the specified shape in the z-order, the very bottom shape's z-order value is 0. Read-only.|beta|
|[shape](reference/excel/shape.md)|_Relationship_ > fill|Returns the fill formatting of the shape object. Read-only.|beta|
|[shape](reference/excel/shape.md)|_Relationship_ > geometricShape|Returns the geometric shape for the shape object. Error will be thrown, if the shape object is other shape type (Like, Image, SmartArt, etc.) rather than GeometricShape. Read-only.|beta|
|[shape](reference/excel/shape.md)|_Relationship_ > geometricShapeType|Represents the geometric shape type of the specified shape.|beta|
|[shape](reference/excel/shape.md)|_Relationship_ > image|Returns the image for the shape object. Error will be thrown, if the shape object is other shape type (Like, GeometricShape, SmartArt, etc.) rather than Image. Read-only.|beta|
|[shape](reference/excel/shape.md)|_Relationship_ > placement|Represents the placment, value that represents the way the object is attached to the cells below it.|beta|
|[shape](reference/excel/shape.md)|_Relationship_ > textFrame|Returns the textFrame object of a shape. Read only. Read-only.|beta|
|[shape](reference/excel/shape.md)|_Relationship_ > type|Returns the type of the specified shape. Read-only.|beta|
|[shape](reference/excel/shape.md)|_Method_ > [delete()]((reference/excel/shape.md#delete)|Deletes the Shape|beta|
|[shape](reference/excel/shape.md)|_Method_ > [setZOrder(value: string)]((reference/excel/shape.md#setzordervalue-string)|Moves the specified shape in front of or behind other shapes in the collection (that is, changes the shape's position in the z-order).|beta|
|[shapeActivatedEventArgs](reference/excel/shapeactivatedeventargs.md)|_Property_ > shapeId|Gets the id of the shape that is activated.|beta|
|[shapeActivatedEventArgs](reference/excel/shapeactivatedeventargs.md)|_Property_ > type|Gets the type of the event.|beta|
|[shapeActivatedEventArgs](reference/excel/shapeactivatedeventargs.md)|_Property_ > worksheetId|Gets the id of the worksheet in which the shape is activated.|beta|
|[shapeCollection](reference/excel/shapecollection.md)|_Property_ > items|A collection of shape objects. Read-only.|beta|
|[shapeCollection](reference/excel/shapecollection.md)|_Method_ > [addGeometricShape(geometricShapeType: string, left: double, top: double, width: double, height: double)]((reference/excel/shapecollection.md#addgeometricshapegeometricshapetype-string-left-double-top-double-width-double-height-double)|Adds a geometric shape to worksheet. Returns a Shape object that represents the new shape.|beta|
|[shapeCollection](reference/excel/shapecollection.md)|_Method_ > [addImage(base64ImageString: string)]((reference/excel/shapecollection.md#addimagebase64imagestring-string)|Creates an image from a base64 string and adds it to worksheet. Returns the image object that represents the new Image.|beta|
|[shapeCollection](reference/excel/shapecollection.md)|_Method_ > [addTextBox(text: string)]((reference/excel/shapecollection.md#addtextboxtext-string)|Adds a textbox to worksheet by telling it's text content. Returns a Shape object that represents the new text box.|beta|
|[shapeCollection](reference/excel/shapecollection.md)|_Method_ > [getCount()]((reference/excel/shapecollection.md#getcount)|Returns the number of shapes in the worksheet. Read-only.|beta|
|[shapeCollection](reference/excel/shapecollection.md)|_Method_ > [getItem(shapeId: string)]((reference/excel/shapecollection.md#getitemshapeid-string)|Returns a shape identified by the shape id. Read-only.|beta|
|[shapeDeactivatedEventArgs](reference/excel/shapedeactivatedeventargs.md)|_Property_ > shapeId|Gets the id of the shape that is deactivated.|beta|
|[shapeDeactivatedEventArgs](reference/excel/shapedeactivatedeventargs.md)|_Property_ > type|Gets the type of the event.|beta|
|[shapeDeactivatedEventArgs](reference/excel/shapedeactivatedeventargs.md)|_Property_ > worksheetId|Gets the id of the worksheet in which the shape is deactivated.|beta|
|[shapeFill](reference/excel/shapefill.md)|_Property_ > foreColor|Represents the shape fill fore color in HTML color format, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange")|beta|
|[shapeFill](reference/excel/shapefill.md)|_Property_ > transparency|Returns or sets the degree of transparency of the specified fill as a value from 0.0 (opaque) through 1.0 (clear). For API not supported shape types  or special fill type with inconsistent transparencies, return null. For example, gradient fill type could have inconsistent transparencies.|beta|
|[shapeFill](reference/excel/shapefill.md)|_Relationship_ > type|Returns the fill type of the shape. Read-only.|beta|
|[shapeFill](reference/excel/shapefill.md)|_Method_ > [clear()]((reference/excel/shapefill.md#clear)|Clears the fill formatting of a shape object.|beta|
|[shapeFill](reference/excel/shapefill.md)|_Method_ > [setSolidColor(color: string)]((reference/excel/shapefill.md#setsolidcolorcolor-string)|Sets the fill formatting of a shape object to a uniform color, fill type changeing to Solid Fill.|beta|
|[shapeFont](reference/excel/shapefont.md)|_Property_ > bold|Represents the bold status of font. Returns null the TextRange includes both bold and non-bold text fragments.|beta|
|[shapeFont](reference/excel/shapefont.md)|_Property_ > color|HTML color code representation of the text color. E.g. #FF0000 represents Red. Returns null if the TextRange includes text fragments with different colors.|beta|
|[shapeFont](reference/excel/shapefont.md)|_Property_ > italic|Represents the italic status of font. Return null if the TextRange includes both italic and non-italic text fragments.|beta|
|[shapeFont](reference/excel/shapefont.md)|_Property_ > name|Represents font name (e.g. "Calibri"). If the text is Complex Script or East Asian language, represents corresponding font name; otherwise represents Latin font name.|beta|
|[shapeFont](reference/excel/shapefont.md)|_Property_ > size|Represents font size in points (e.g. 11). Return null if the TextRange includes text fragments with different font sizes.|beta|
|[shapeFont](reference/excel/shapefont.md)|_Relationship_ > underline|Type of underline applied to the font. Return null if the TextRange includes text fragments with different underline styles.|beta|
|[showAsRule](reference/excel/showasrule.md)|_Relationship_ > baseField|The Base PivotField to base the ShowAs calculation, if applicable based on the ShowAsCalculation type, else null.|1.8|
|[showAsRule](reference/excel/showasrule.md)|_Relationship_ > baseItem|The Base Item to base the ShowAs calculation on, if applicable based on the ShowAsCalculation type, else null.|1.8|
|[showAsRule](reference/excel/showasrule.md)|_Relationship_ > calculation|The ShowAs Calculation to use for the Data PivotField.|1.8|
|[showCardPostProcessAction](reference/excel/showcardpostprocessaction.md)|_Property_ > column|Transmits additional data to client side, e.g., worksheetId for TableSelectionChangedEvent.|1.7|
|[showCardPostProcessAction](reference/excel/showcardpostprocessaction.md)|_Property_ > row|Transmits additional data to client side, e.g., worksheetId for TableSelectionChangedEvent.|1.7|
|[showCardPostProcessAction](reference/excel/showcardpostprocessaction.md)|_Relationship_ > actionType|Transmits additional data to client side, e.g., worksheetId for TableSelectionChangedEvent.|1.7|
|[sortField](reference/excel/sortfield.md)|_Property_ > subField|Represents the subfield that is the target property name of a rich value to sort on.|beta|
|[style](reference/excel/style.md)|_Property_ > addIndent|Indicates if text is automatically indented when the text alignment in a cell is set to equal distribution.|1.8|
|[style](reference/excel/style.md)|_Property_ > autoIndent|Indicates if text is automatically indented when the text alignment in a cell is set to equal distribution.|1.8|
|[style](reference/excel/style.md)|_Property_ > builtIn|Indicates if the style is a built-in style. Read-only.|1.7|
|[style](reference/excel/style.md)|_Property_ > formulaHidden|Indicates if the formula will be hidden when the worksheet is protected.|1.7|
|[style](reference/excel/style.md)|_Property_ > horizontalAlignment|Represents the horizontal alignment for the style. Possible values are: General, Left, Center, Right, Fill, Justify, CenterAcrossSelection, Distributed.|1.7|
|[style](reference/excel/style.md)|_Property_ > includeAlignment|Indicates if the style includes the AutoIndent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel, and TextOrientation properties.|1.7|
|[style](reference/excel/style.md)|_Property_ > includeBorder|Indicates if the style includes the Color, ColorIndex, LineStyle, and Weight border properties.|1.7|
|[style](reference/excel/style.md)|_Property_ > includeFont|Indicates if the style includes the Background, Bold, Color, ColorIndex, FontStyle, Italic, Name, Size, Strikethrough, Subscript, Superscript, and Underline font properties.|1.7|
|[style](reference/excel/style.md)|_Property_ > includeNumber|Indicates if the style includes the NumberFormat property.|1.7|
|[style](reference/excel/style.md)|_Property_ > includePatterns|Indicates if the style includes the Color, ColorIndex, InvertIfNegative, Pattern, PatternColor, and PatternColorIndex interior properties.|1.7|
|[style](reference/excel/style.md)|_Property_ > includeProtection|Indicates if the style includes the FormulaHidden and Locked protection properties.|1.7|
|[style](reference/excel/style.md)|_Property_ > indentLevel|An integer from 0 to 250 that indicates the indent level for the style.|1.7|
|[style](reference/excel/style.md)|_Property_ > locked|Indicates if the object is locked when the worksheet is protected.|1.7|
|[style](reference/excel/style.md)|_Property_ > name|The name of the style. Read-only.|1.7|
|[style](reference/excel/style.md)|_Property_ > numberFormat|The format code of the number format for the style.|1.7|
|[style](reference/excel/style.md)|_Property_ > numberFormatLocal|The localized format code of the number format for the style.|1.7|
|[style](reference/excel/style.md)|_Property_ > orientation|The text orientation for the style.|1.8|
|[style](reference/excel/style.md)|_Property_ > readingOrder|The reading order for the style. Possible values are: Context, LeftToRight, RightToLeft.|1.7|
|[style](reference/excel/style.md)|_Property_ > shrinkToFit|Indicates if text automatically shrinks to fit in the available column width.|1.7|
|[style](reference/excel/style.md)|_Property_ > textOrientation|The text orientation for the style.|1.8|
|[style](reference/excel/style.md)|_Property_ > verticalAlignment|Represents the vertical alignment for the style. Possible values are: Top, Center, Bottom, Justify, Distributed.|1.7|
|[style](reference/excel/style.md)|_Property_ > wrapText|Indicates if Microsoft Excel wraps the text in the object.|1.7|
|[style](reference/excel/style.md)|_Relationship_ > borders|A Border collection of four Border objects that represent the style of the four borders. Read-only.|1.7|
|[style](reference/excel/style.md)|_Relationship_ > fill|The Fill of the style. Read-only.|1.7|
|[style](reference/excel/style.md)|_Relationship_ > font|A Font object that represents the font of the style. Read-only.|1.7|
|[style](reference/excel/style.md)|_Method_ > [delete()]((reference/excel/style.md#delete)|Deletes this style.|1.7|
|[styleCollection](reference/excel/stylecollection.md)|_Property_ > items|A collection of style objects. Read-only.|1.7|
|[styleCollection](reference/excel/stylecollection.md)|_Method_ > [add(name: string)]((reference/excel/stylecollection.md#addname-string)|Adds a new style to the collection.|1.7|
|[styleCollection](reference/excel/stylecollection.md)|_Method_ > [getItem(name: string)]((reference/excel/stylecollection.md#getitemname-string)|Gets a style by name.|1.7|
|[subtotals](reference/excel/subtotals.md)|_Property_ > automatic|If Automatic is set to true, then all other values will be ignored when setting the Subtotals.|1.8|
|[subtotals](reference/excel/subtotals.md)|_Property_ > average|If Automatic is set to true, then all other values will be ignored when setting the Subtotals.|1.8|
|[subtotals](reference/excel/subtotals.md)|_Property_ > count|If Automatic is set to true, then all other values will be ignored when setting the Subtotals.|1.8|
|[subtotals](reference/excel/subtotals.md)|_Property_ > countNumbers|If Automatic is set to true, then all other values will be ignored when setting the Subtotals.|1.8|
|[subtotals](reference/excel/subtotals.md)|_Property_ > max|If Automatic is set to true, then all other values will be ignored when setting the Subtotals.|1.8|
|[subtotals](reference/excel/subtotals.md)|_Property_ > min|If Automatic is set to true, then all other values will be ignored when setting the Subtotals.|1.8|
|[subtotals](reference/excel/subtotals.md)|_Property_ > product|If Automatic is set to true, then all other values will be ignored when setting the Subtotals.|1.8|
|[subtotals](reference/excel/subtotals.md)|_Property_ > standardDeviation|If Automatic is set to true, then all other values will be ignored when setting the Subtotals.|1.8|
|[subtotals](reference/excel/subtotals.md)|_Property_ > standardDeviationP|If Automatic is set to true, then all other values will be ignored when setting the Subtotals.|1.8|
|[subtotals](reference/excel/subtotals.md)|_Property_ > sum|If Automatic is set to true, then all other values will be ignored when setting the Subtotals.|1.8|
|[subtotals](reference/excel/subtotals.md)|_Property_ > variance|If Automatic is set to true, then all other values will be ignored when setting the Subtotals.|1.8|
|[subtotals](reference/excel/subtotals.md)|_Property_ > varianceP|If Automatic is set to true, then all other values will be ignored when setting the Subtotals.|1.8|
|[table](reference/excel/table.md)|_Property_ > legacyId|Returns a numeric id. Read-only.|1.8|
|[table](reference/excel/table.md)|_Relationship_ > autoFilter|Represents the AutoFilter object of the table. Read-Only. Read-only.|beta|
|[table](reference/excel/table.md)|_Method_ > [clearStyle()]((reference/excel/table.md#clearstyle)|Changes the table to use the default table style.|beta|
|[tableAddedEventArgs](reference/excel/tableaddedeventargs.md)|_Property_ > tableId|Gets the id of the table that is added.|beta|
|[tableAddedEventArgs](reference/excel/tableaddedeventargs.md)|_Property_ > type|Gets the type of the event.|beta|
|[tableAddedEventArgs](reference/excel/tableaddedeventargs.md)|_Property_ > worksheetId|Gets the id of the worksheet in which the table is added.|beta|
|[tableAddedEventArgs](reference/excel/tableaddedeventargs.md)|_Relationship_ > source|Gets the source of the event.|beta|
|[tableChangedEventArgs](reference/excel/tablechangedeventargs.md)|_Property_ > address|Gets the address that represents the changed area of a table on a specific worksheet.|1.7|
|[tableChangedEventArgs](reference/excel/tablechangedeventargs.md)|_Property_ > tableId|Gets the id of the table in which the data changed.|1.7|
|[tableChangedEventArgs](reference/excel/tablechangedeventargs.md)|_Property_ > type|Gets the type of the event.|1.7|
|[tableChangedEventArgs](reference/excel/tablechangedeventargs.md)|_Property_ > worksheetId|Gets the id of the worksheet in which the data changed.|1.7|
|[tableChangedEventArgs](reference/excel/tablechangedeventargs.md)|_Relationship_ > changeType|Gets the change type that represents how the Changed event is triggered.|1.7|
|[tableChangedEventArgs](reference/excel/tablechangedeventargs.md)|_Relationship_ > source|Gets the source of the event.|1.7|
|[tableColumn](reference/excel/tablecolumn.md)|_Method_ > [getNextColumn()]((reference/excel/tablecolumn.md#getnextcolumn)|Gets the table column that follows this one. If there are no table columns following this one, this method will throw an error.|ApiSetAttribute.Spec|
|[tableColumn](reference/excel/tablecolumn.md)|_Method_ > [getNextColumnOrNullObject()]((reference/excel/tablecolumn.md#getnextcolumnornullobject)|Gets the table column that follows this one. If there are no table columns following this one, this method will return a null object.|ApiSetAttribute.Spec|
|[tableColumn](reference/excel/tablecolumn.md)|_Method_ > [getPreviousColumn()]((reference/excel/tablecolumn.md#getpreviouscolumn)|Gets the table column that precedes this one. If there are no previous table columns, this method will throw an error.|ApiSetAttribute.Spec|
|[tableColumn](reference/excel/tablecolumn.md)|_Method_ > [getPreviousColumnOrNullObject()]((reference/excel/tablecolumn.md#getpreviouscolumnornullobject)|Gets the table column that precedes this one. If there are no previous table columns, this method will return a null objet.|ApiSetAttribute.Spec|
|[tableColumnCollection](reference/excel/tablecolumncollection.md)|_Method_ > [getFirst()]((reference/excel/tablecolumncollection.md#getfirst)|Gets the first column in the table.|ApiSetAttribute.Spec|
|[tableColumnCollection](reference/excel/tablecolumncollection.md)|_Method_ > [getLast()]((reference/excel/tablecolumncollection.md#getlast)|Gets the last column in the table.|ApiSetAttribute.Spec|
|[tableDeletedEventArgs](reference/excel/tabledeletedeventargs.md)|_Property_ > tableId|Specifies the id of the table that is deleted.|beta|
|[tableDeletedEventArgs](reference/excel/tabledeletedeventargs.md)|_Property_ > tableName|Specifies the name of the table that is deleted.|beta|
|[tableDeletedEventArgs](reference/excel/tabledeletedeventargs.md)|_Property_ > type|Specifies the type of the event.|beta|
|[tableDeletedEventArgs](reference/excel/tabledeletedeventargs.md)|_Property_ > worksheetId|Specifies the id of the worksheet in which the table is deleted.|beta|
|[tableDeletedEventArgs](reference/excel/tabledeletedeventargs.md)|_Relationship_ > source|Specifies the source of the event.|beta|
|[tableFilteredEventArgs](reference/excel/tablefilteredeventargs.md)|_Property_ > tableId|Represents the id of the table in which the filter is applied..|beta|
|[tableFilteredEventArgs](reference/excel/tablefilteredeventargs.md)|_Property_ > type|Represents the type of the event.|beta|
|[tableFilteredEventArgs](reference/excel/tablefilteredeventargs.md)|_Property_ > worksheetId|Represents the id of the worksheet which contains the table.|beta|
|[tableScopedCollection](reference/excel/tablescopedcollection.md)|_Property_ > items|A collection of tableScoped objects. Read-only.|beta|
|[tableScopedCollection](reference/excel/tablescopedcollection.md)|_Method_ > [getCount()]((reference/excel/tablescopedcollection.md#getcount)|Gets the number of tables in the collection.|beta|
|[tableScopedCollection](reference/excel/tablescopedcollection.md)|_Method_ > [getFirst()]((reference/excel/tablescopedcollection.md#getfirst)|Gets the first table in the collection. The tables in the collection are sorted top to bottom and left to right, such that top left table is the first table in the collection.|beta|
|[tableScopedCollection](reference/excel/tablescopedcollection.md)|_Method_ > [getItem(key: string)]((reference/excel/tablescopedcollection.md#getitemkey-string)|Gets a table by Name or ID.|beta|
|[tableSelectionChangedEventArgs](reference/excel/tableselectionchangedeventargs.md)|_Property_ > address|Gets the range address that represents the selected area of the table on a specific worksheet.|1.7|
|[tableSelectionChangedEventArgs](reference/excel/tableselectionchangedeventargs.md)|_Property_ > isInsideTable|Indicates if the selection is inside a table, address will be useless if IsInsideTable is false.|1.7|
|[tableSelectionChangedEventArgs](reference/excel/tableselectionchangedeventargs.md)|_Property_ > tableId|Gets the id of the table in which the selection changed.|1.7|
|[tableSelectionChangedEventArgs](reference/excel/tableselectionchangedeventargs.md)|_Property_ > type|Gets the type of the event.|1.7|
|[tableSelectionChangedEventArgs](reference/excel/tableselectionchangedeventargs.md)|_Property_ > worksheetId|Gets the id of the worksheet in which the selection changed.|1.7|
|[textFrame](reference/excel/textframe.md)|_Property_ > bottomMargin|Represents the bottom margin, in points, of the text frame.|beta|
|[textFrame](reference/excel/textframe.md)|_Property_ > hasText|Specifies whether the TextFrame contains text. Read-only.|beta|
|[textFrame](reference/excel/textframe.md)|_Property_ > leftMargin|Represents the left margin, in points, of the text frame.|beta|
|[textFrame](reference/excel/textframe.md)|_Property_ > rightMargin|Represents the right margin, in points, of the text frame.|beta|
|[textFrame](reference/excel/textframe.md)|_Property_ > topMargin|Represents the top margin, in points, of the text frame.|beta|
|[textFrame](reference/excel/textframe.md)|_Relationship_ > autoSize|Gets or sets the auto sizing settings for the text frame. A text frame can be set to auto size the text to fit the text frame, or auto size the text frame to fit the text, or without auto sizing.|beta|
|[textFrame](reference/excel/textframe.md)|_Relationship_ > horizontalAlignment|Represents the horizontal alignment of the text frame.|beta|
|[textFrame](reference/excel/textframe.md)|_Relationship_ > horizontalOverflow|Represents the horizontal overflow type of the text frame.|beta|
|[textFrame](reference/excel/textframe.md)|_Relationship_ > orientation|Represents the text orientation of the text frame.|beta|
|[textFrame](reference/excel/textframe.md)|_Relationship_ > readingOrder|Represents the reading order of the text frame, RTL or LTR.|beta|
|[textFrame](reference/excel/textframe.md)|_Relationship_ > textRange|Returns or sets the degree of transparency of the specified fill as a value from 0.0 (opaque) through 1.0 (clear). For API not supported shape types  or special fill type with inconsistent transparencies, return null. For example, gradient fill type could have inconsistent transparencies. Read-only.|beta|
|[textFrame](reference/excel/textframe.md)|_Relationship_ > verticalAlignment|Represents the vertical alignment of the text frame.|beta|
|[textFrame](reference/excel/textframe.md)|_Relationship_ > verticalOverflow|Represents the vertical overflow type of the text frame.|beta|
|[textFrame](reference/excel/textframe.md)|_Method_ > [deleteText()]((reference/excel/textframe.md#deletetext)|Deletes all the text in the textframe.|beta|
|[textRange](reference/excel/textrange.md)|_Property_ > text|Represents the plain text content of the text range.|beta|
|[textRange](reference/excel/textrange.md)|_Relationship_ > font|Returns a ShapeFont object that represents the font attributes for the text range. Read-only.|beta|
|[textRange](reference/excel/textrange.md)|_Method_ > [getCharacters(start: number, length: number)]((reference/excel/textrange.md#getcharactersstart-number-length-number)|Returns a TextRange object for characters in the given range.|beta|
|[workbook](reference/excel/workbook.md)|_Property_ > autoSave|True if the workbook is in auto save mode. Read-only.|beta|
|[workbook](reference/excel/workbook.md)|_Property_ > calculationEngineVersion|Returns a number about the version of Excel Calculation Engine. Read-Only. Read-only.|beta|
|[workbook](reference/excel/workbook.md)|_Property_ > chartDataPointTrack|True if all charts in the workbook are tracking the actual data points to which they are attached.|beta|
|[workbook](reference/excel/workbook.md)|_Property_ > isDirty|True if no changes have been made to the specified workbook since it was last saved.|beta|
|[workbook](reference/excel/workbook.md)|_Property_ > name|Gets the workbook name. Read-only.|1.7|
|[workbook](reference/excel/workbook.md)|_Property_ > previouslySaved|True if the workbook has ever been saved locally or online. Read-only.|beta|
|[workbook](reference/excel/workbook.md)|_Property_ > readOnly|True if the workbook is open in Read-only mode. Read-only.|1.8|
|[workbook](reference/excel/workbook.md)|_Property_ > use1904DateSystem|True if the workbook uses the 1904 date system.|beta|
|[workbook](reference/excel/workbook.md)|_Property_ > usePrecisionAsDisplayed|True if calculations in this workbook will be done using only the precision of the numbers as they're displayed.|beta|
|[workbook](reference/excel/workbook.md)|_Relationship_ > dataConnections|Represents all data connections in the workbook. Read-only.|1.7|
|[workbook](reference/excel/workbook.md)|_Relationship_ > properties|Gets the workbook properties. Read-only.|1.7|
|[workbook](reference/excel/workbook.md)|_Relationship_ > protection|Returns workbook protection object for a workbook. Read-only.|1.7|
|[workbook](reference/excel/workbook.md)|_Relationship_ > styles|Represents a collection of styles associated with the workbook. Read-only.|1.7|
|[workbook](reference/excel/workbook.md)|_Method_ > [close(closeBehavior: CloseBehavior)]((reference/excel/workbook.md#closeclosebehavior-closebehavior)|Close current workbook.|beta|
|[workbook](reference/excel/workbook.md)|_Method_ > [getActiveCell()]((reference/excel/workbook.md#getactivecell)|Gets the currently active cell from the workbook.|1.7|
|[workbook](reference/excel/workbook.md)|_Method_ > [getActiveChart()]((reference/excel/workbook.md#getactivechart)|Gets the currently active chart in the workbook. If there is no active chart, will throw exception when invoke this statement|beta|
|[workbook](reference/excel/workbook.md)|_Method_ > [getActiveChartOrNullObject()]((reference/excel/workbook.md#getactivechartornullobject)|Gets the currently active chart in the workbook. If there is no active chart, will return null object|beta|
|[workbook](reference/excel/workbook.md)|_Method_ > [getIsActiveCollabSession()]((reference/excel/workbook.md#getisactivecollabsession)|True if the workbook is being edited by multiple users (co-authoring).|beta|
|[workbook](reference/excel/workbook.md)|_Method_ > [getRange(address: string)]((reference/excel/workbook.md#getrangeaddress-string)|Gets the range object specified by the address or name.|ApiSetAttribute.Spec|
|[workbook](reference/excel/workbook.md)|_Method_ > [getSelectedRanges()]((reference/excel/workbook.md#getselectedranges)|Gets the currently selected one or more ranges from the workbook. Unlike getSelectedRange(), this method returns a RangeAreas object that represents all the selected ranges.|beta|
|[workbookCreated](reference/excel/workbookcreated.md)|_Property_ > id|Returns a value that uniquely identifies the WorkbookCreated object. Read-only.|1.8|
|[workbookCreated](reference/excel/workbookcreated.md)|_Method_ > [open()]((reference/excel/workbookcreated.md#open)|Open the workbook.|1.8|
|[workbookProtection](reference/excel/workbookprotection.md)|_Property_ > protected|Indicates if the workbook is protected. Read-Only. Read-only.|1.7|
|[workbookProtection](reference/excel/workbookprotection.md)|_Method_ > [protect(password: string)]((reference/excel/workbookprotection.md#protectpassword-string)|Protects a workbook. Fails if the workbook has been protected.|1.7|
|[workbookProtection](reference/excel/workbookprotection.md)|_Method_ > [unprotect(password: string)]((reference/excel/workbookprotection.md#unprotectpassword-string)|Unprotects a workbook.|1.7|
|[worksheet](reference/excel/worksheet.md)|_Property_ > enableCalculation|Gets or sets the enableCalculation property of the worksheet.|beta|
|[worksheet](reference/excel/worksheet.md)|_Property_ > gridlines|Gets or sets the worksheet's gridlines flag.|1.8|
|[worksheet](reference/excel/worksheet.md)|_Property_ > headings|Gets or sets the worksheet's headings flag.|1.8|
|[worksheet](reference/excel/worksheet.md)|_Property_ > showGridlines|Gets or sets the worksheet's gridlines flag.|1.8|
|[worksheet](reference/excel/worksheet.md)|_Property_ > showHeadings|Gets or sets the worksheet's headings flag.|1.8|
|[worksheet](reference/excel/worksheet.md)|_Property_ > standardHeight|Returns the standard (default) height of all the rows in the worksheet, in points. Read-only.|1.7|
|[worksheet](reference/excel/worksheet.md)|_Property_ > standardWidth|Returns or sets the standard (default) width of all the columns in the worksheet.|1.7|
|[worksheet](reference/excel/worksheet.md)|_Property_ > tabColor|Gets or sets the worksheet tab color.|1.7|
|[worksheet](reference/excel/worksheet.md)|_Relationship_ > autoFilter|Represents the AutoFilter object of the worksheet. Read-Only. Read-only.|beta|
|[worksheet](reference/excel/worksheet.md)|_Relationship_ > freezePanes|Gets an object that can be used to manipulate frozen panes on the worksheet. Read-only.|1.7|
|[worksheet](reference/excel/worksheet.md)|_Relationship_ > horizontalPageBreaks|Gets the horizontal page break collection for the worksheet. This collection only contains manual page breaks. Read-only.|beta|
|[worksheet](reference/excel/worksheet.md)|_Relationship_ > pageLayout|Gets the PageLayout object of the worksheet. Read-only.|beta|
|[worksheet](reference/excel/worksheet.md)|_Relationship_ > shapes|Returns the collection of all the Shape objects on the worksheet. Read-only.|beta|
|[worksheet](reference/excel/worksheet.md)|_Relationship_ > verticalPageBreaks|Gets the vertical page break collection for the worksheet. This collection only contains manual page breaks. Read-only.|beta|
|[worksheet](reference/excel/worksheet.md)|_Method_ > [copy(positionType: WorksheetPositionType, relativeTo: Worksheet)]((reference/excel/worksheet.md#copypositiontype-worksheetpositiontype-relativeto-worksheet)|Copy a worksheet and place it at the specified position. Return the copied worksheet.|1.7|
|[worksheet](reference/excel/worksheet.md)|_Method_ > [findAll(text: string, criteria: WorksheetSearchCriteria)]((reference/excel/worksheet.md#findalltext-string-criteria-worksheetsearchcriteria)|Finds all occurrences of the given string based on the criteria specified and returns them as a RangeAreas object, comprising one or more rectangular ranges.|beta|
|[worksheet](reference/excel/worksheet.md)|_Method_ > [findAllOrNullObject(text: string, criteria: WorksheetSearchCriteria)]((reference/excel/worksheet.md#findallornullobjecttext-string-criteria-worksheetsearchcriteria)|Finds all occurrences of the given string based on the criteria specified and returns them as a RangeAreas object, comprising one or more rectangular ranges.|beta|
|[worksheet](reference/excel/worksheet.md)|_Method_ > [getRangeByIndexes(startRow: number, startColumn: number, rowCount: number, columnCount: number)]((reference/excel/worksheet.md#getrangebyindexesstartrow-number-startcolumn-number-rowcount-number-columncount-number)|Gets the range object beginning at a particular row index and column index, and spanning a certain number of rows and columns.|1.7|
|[worksheet](reference/excel/worksheet.md)|_Method_ > [getRanges(address: string)]((reference/excel/worksheet.md#getrangesaddress-string)|Gets the RangeAreas object, representing one or more blocks of rectangular ranges, specified by the address or name.|beta|
|[worksheet](reference/excel/worksheet.md)|_Method_ > [replaceAll(text: string, replacement: string, criteria: ReplaceCriteria)]((reference/excel/worksheet.md#replacealltext-string-replacement-string-criteria-replacecriteria)|Finds and replaces the given string based on the criteria specified within the current worksheet.|beta|
|[worksheetActivatedEventArgs](reference/excel/worksheetactivatedeventargs.md)|_Property_ > type|Gets the type of the event.|1.7|
|[worksheetActivatedEventArgs](reference/excel/worksheetactivatedeventargs.md)|_Property_ > worksheetId|Gets the id of the worksheet that is activated.|1.7|
|[worksheetAddedEventArgs](reference/excel/worksheetaddedeventargs.md)|_Property_ > type|Gets the type of the event.|1.7|
|[worksheetAddedEventArgs](reference/excel/worksheetaddedeventargs.md)|_Property_ > worksheetId|Gets the id of the worksheet that is added to the workbook.|1.7|
|[worksheetAddedEventArgs](reference/excel/worksheetaddedeventargs.md)|_Relationship_ > source|Gets the source of the event.|1.7|
|[worksheetCalculatedEventArgs](reference/excel/worksheetcalculatedeventargs.md)|_Property_ > type|Gets the type of the event.|1.8|
|[worksheetCalculatedEventArgs](reference/excel/worksheetcalculatedeventargs.md)|_Property_ > worksheetId|Gets the id of the worksheet that is calculated.|1.8|
|[worksheetChangedEventArgs](reference/excel/worksheetchangedeventargs.md)|_Property_ > address|Gets the range address that represents the changed area of a specific worksheet.|1.7|
|[worksheetChangedEventArgs](reference/excel/worksheetchangedeventargs.md)|_Property_ > type|Gets the type of the event.|1.7|
|[worksheetChangedEventArgs](reference/excel/worksheetchangedeventargs.md)|_Property_ > worksheetId|Gets the id of the worksheet in which the data changed.|1.7|
|[worksheetChangedEventArgs](reference/excel/worksheetchangedeventargs.md)|_Relationship_ > changeType|Gets the change type that represents how the Changed event is triggered.|1.7|
|[worksheetChangedEventArgs](reference/excel/worksheetchangedeventargs.md)|_Relationship_ > source|Gets the source of the event.|1.7|
|[worksheetDeactivatedEventArgs](reference/excel/worksheetdeactivatedeventargs.md)|_Property_ > type|Gets the type of the event.|1.7|
|[worksheetDeactivatedEventArgs](reference/excel/worksheetdeactivatedeventargs.md)|_Property_ > worksheetId|Gets the id of the worksheet that is deactivated.|1.7|
|[worksheetDeletedEventArgs](reference/excel/worksheetdeletedeventargs.md)|_Property_ > type|Gets the type of the event.|1.7|
|[worksheetDeletedEventArgs](reference/excel/worksheetdeletedeventargs.md)|_Property_ > worksheetId|Gets the id of the worksheet that is deleted from the workbook.|1.7|
|[worksheetDeletedEventArgs](reference/excel/worksheetdeletedeventargs.md)|_Relationship_ > source|Gets the source of the event.|1.7|
|[worksheetFilteredEventArgs](reference/excel/worksheetfilteredeventargs.md)|_Property_ > type|Represents the type of the event.|beta|
|[worksheetFilteredEventArgs](reference/excel/worksheetfilteredeventargs.md)|_Property_ > worksheetId|Represents the id of the worksheet in which the filter is applied.|beta|
|[worksheetFreezePanes](reference/excel/worksheetfreezepanes.md)|_Method_ > [freezeAt(frozenRange: Range or string)]((reference/excel/worksheetfreezepanes.md#freezeatfrozenrange-range-or-string)|Sets the frozen cells in the active worksheet view.|1.7|
|[worksheetFreezePanes](reference/excel/worksheetfreezepanes.md)|_Method_ > [freezeColumns(count: number)]((reference/excel/worksheetfreezepanes.md#freezecolumnscount-number)|Freeze the first column(s) of the worksheet in place.|1.7|
|[worksheetFreezePanes](reference/excel/worksheetfreezepanes.md)|_Method_ > [freezeRows(count: number)]((reference/excel/worksheetfreezepanes.md#freezerowscount-number)|Freeze the top row(s) of the worksheet in place.|1.7|
|[worksheetFreezePanes](reference/excel/worksheetfreezepanes.md)|_Method_ > [getLocation()]((reference/excel/worksheetfreezepanes.md#getlocation)|Gets a range that describes the frozen cells in the active worksheet view.|1.7|
|[worksheetFreezePanes](reference/excel/worksheetfreezepanes.md)|_Method_ > [getLocationOrNullObject()]((reference/excel/worksheetfreezepanes.md#getlocationornullobject)|Gets a range that describes the frozen cells in the active worksheet view.|1.7|
|[worksheetFreezePanes](reference/excel/worksheetfreezepanes.md)|_Method_ > [unfreeze()]((reference/excel/worksheetfreezepanes.md#unfreeze)|Removes all frozen panes in the worksheet.|1.7|
|[worksheetProtectionOptions](reference/excel/worksheetprotectionoptions.md)|_Property_ > allowEditObjects|Represents the worksheet protection option of allowing editing objects.|1.7|
|[worksheetProtectionOptions](reference/excel/worksheetprotectionoptions.md)|_Property_ > allowEditScenarios|Represents the worksheet protection option of allowing editing scenarios.|1.7|
|[worksheetProtectionOptions](reference/excel/worksheetprotectionoptions.md)|_Relationship_ > selectionMode|Represents the worksheet protection option of selection mode.|1.7|
|[worksheetSearchCriteria](reference/excel/worksheetsearchcriteria.md)|_Property_ > completeMatch|Specifies whether the match needs to be complete or partial. Default is false (partial).|beta|
|[worksheetSearchCriteria](reference/excel/worksheetsearchcriteria.md)|_Property_ > matchCase|Specifies whether the match is case sensitive. Default is false (insensitive).|beta|
|[worksheetSelectionChangedEventArgs](reference/excel/worksheetselectionchangedeventargs.md)|_Property_ > address|Gets the range address that represents the selected area of a specific worksheet.|1.7|
|[worksheetSelectionChangedEventArgs](reference/excel/worksheetselectionchangedeventargs.md)|_Property_ > type|Gets the type of the event.|1.7|
|[worksheetSelectionChangedEventArgs](reference/excel/worksheetselectionchangedeventargs.md)|_Property_ > worksheetId|Gets the id of the worksheet in which the selection changed.|1.7|
|[_ChartAddedEventArgs](reference/excel/_chartaddedeventargs.md)|_Property_ > chartId|Gets the id of the worksheet that is deleted from the workbook.|1.8|
|[_ChartAddedEventArgs](reference/excel/_chartaddedeventargs.md)|_Property_ > worksheetId|Gets the id of the worksheet that is deleted from the workbook.|1.8|
|[_ChartAddedEventArgs](reference/excel/_chartaddedeventargs.md)|_Relationship_ > source|Gets the id of the worksheet that is deleted from the workbook.|1.8|
|[_ChartDeletedEventArgs](reference/excel/_chartdeletedeventargs.md)|_Property_ > chartId|Gets the id of the chart that is deactivated.|1.8|
|[_ChartDeletedEventArgs](reference/excel/_chartdeletedeventargs.md)|_Property_ > worksheetId|Gets the id of the chart that is deactivated.|1.8|
|[_ChartDeletedEventArgs](reference/excel/_chartdeletedeventargs.md)|_Relationship_ > source|Gets the id of the chart that is deactivated.|1.8|
|[_InternalTestEventArgs](reference/excel/_internaltesteventargs.md)|_Property_ > prop1|Gets a style by name.|1.7|
|[_InternalTestEventArgs](reference/excel/_internaltesteventargs.md)|_Property_ > worksheetId|Gets a style by name.|1.7|
|[_MessageEventArgs](reference/excel/_messageeventargs.md)|_Relationship_ > workbook|Gets the Workbook object that raised the Message event|1.7|
|[_MessageEventArgsEntry](reference/excel/_messageeventargsentry.md)|_Property_ > isRemoteOverride|Gets the Workbook object that raised the Message event|1.7|
|[_MessageEventArgsEntry](reference/excel/_messageeventargsentry.md)|_Property_ > message|Gets the Workbook object that raised the Message event|1.7|
|[_MessageEventArgsEntry](reference/excel/_messageeventargsentry.md)|_Property_ > targetId|Gets the Workbook object that raised the Message event|1.7|
|[_MessageEventArgsEntry](reference/excel/_messageeventargsentry.md)|_Relationship_ > messageCategory|Gets the Workbook object that raised the Message event|1.7|
|[_MessageEventArgsEntry](reference/excel/_messageeventargsentry.md)|_Relationship_ > messageType|Gets the Workbook object that raised the Message event|1.7|
|[_TableAddedEventArgs](reference/excel/_tableaddedeventargs.md)|_Property_ > tableId|Gets the id of the worksheet that is calculated.|beta|
|[_TableAddedEventArgs](reference/excel/_tableaddedeventargs.md)|_Property_ > worksheetId|Gets the id of the worksheet that is calculated.|beta|
|[_TableAddedEventArgs](reference/excel/_tableaddedeventargs.md)|_Relationship_ > source|Gets the id of the worksheet that is calculated.|beta|
|[_TableDataChangedEventArgs](reference/excel/_tabledatachangedeventargs.md)|_Property_ > address|Gets the range address that represents the changed area of a specific worksheet.|1.7|
|[_TableDataChangedEventArgs](reference/excel/_tabledatachangedeventargs.md)|_Property_ > referenceId|Gets the range address that represents the changed area of a specific worksheet.|1.8|
|[_TableDataChangedEventArgs](reference/excel/_tabledatachangedeventargs.md)|_Property_ > tableId|Gets the range address that represents the changed area of a specific worksheet.|1.7|
|[_TableDataChangedEventArgs](reference/excel/_tabledatachangedeventargs.md)|_Property_ > worksheetId|Gets the range address that represents the changed area of a specific worksheet.|1.7|
|[_TableDataChangedEventArgs](reference/excel/_tabledatachangedeventargs.md)|_Relationship_ > changeType|Gets the range address that represents the changed area of a specific worksheet.|1.7|
|[_TableDataChangedEventArgs](reference/excel/_tabledatachangedeventargs.md)|_Relationship_ > source|Gets the range address that represents the changed area of a specific worksheet.|1.7|
|[_TableDeletedEventArgs](reference/excel/_tabledeletedeventargs.md)|_Property_ > tableId|Gets the id of the table that is added.|beta|
|[_TableDeletedEventArgs](reference/excel/_tabledeletedeventargs.md)|_Property_ > tableName|Gets the id of the table that is added.|beta|
|[_TableDeletedEventArgs](reference/excel/_tabledeletedeventargs.md)|_Property_ > worksheetId|Gets the id of the table that is added.|beta|
|[_TableDeletedEventArgs](reference/excel/_tabledeletedeventargs.md)|_Relationship_ > source|Gets the id of the table that is added.|beta|
|[_TableFilteredEventArgs](reference/excel/_tablefilteredeventargs.md)|_Property_ > tableId|Gets the address that represents the changed area of a table on a specific worksheet.|beta|
|[_TableFilteredEventArgs](reference/excel/_tablefilteredeventargs.md)|_Property_ > worksheetId|Gets the address that represents the changed area of a table on a specific worksheet.|beta|
|[_TableSelectionChangedEventArgs](reference/excel/_tableselectionchangedeventargs.md)|_Property_ > address|Gets the range address that represents the selected area of the table on a specific worksheet.|1.7|
|[_TableSelectionChangedEventArgs](reference/excel/_tableselectionchangedeventargs.md)|_Property_ > worksheetId|Gets the range address that represents the selected area of the table on a specific worksheet.|1.7|
|[_WorksheetAddedEventArgs](reference/excel/_worksheetaddedeventargs.md)|_Property_ > worksheetId|Gets the range address that represents the selected area of the table on a specific worksheet.|1.7|
|[_WorksheetAddedEventArgs](reference/excel/_worksheetaddedeventargs.md)|_Relationship_ > source|Gets the range address that represents the selected area of the table on a specific worksheet.|1.7|
|[_WorksheetCalculatedEventArgs](reference/excel/_worksheetcalculatedeventargs.md)|_Property_ > worksheetId|Gets the id of the chart that is deleted from the worksheet.|1.8|
|[_WorksheetDataChangedEventArgs](reference/excel/_worksheetdatachangedeventargs.md)|_Property_ > address|Gets the Workbook object that raised the Message event|1.7|
|[_WorksheetDataChangedEventArgs](reference/excel/_worksheetdatachangedeventargs.md)|_Property_ > referenceId|Gets the Workbook object that raised the Message event|1.8|
|[_WorksheetDataChangedEventArgs](reference/excel/_worksheetdatachangedeventargs.md)|_Property_ > worksheetId|Gets the Workbook object that raised the Message event|beta|
|[_WorksheetDataChangedEventArgs](reference/excel/_worksheetdatachangedeventargs.md)|_Relationship_ > changeType|Gets the Workbook object that raised the Message event|1.7|
|[_WorksheetDataChangedEventArgs](reference/excel/_worksheetdatachangedeventargs.md)|_Relationship_ > source|Gets the Workbook object that raised the Message event|1.7|
|[_WorksheetDeletedEventArgs](reference/excel/_worksheetdeletedeventargs.md)|_Property_ > worksheetId|Gets the id of the worksheet that is added to the workbook.|1.8|
|[_WorksheetDeletedEventArgs](reference/excel/_worksheetdeletedeventargs.md)|_Relationship_ > source|Gets the id of the worksheet that is added to the workbook.|1.8|
|[_WorksheetFilteredEventArgs](reference/excel/_worksheetfilteredeventargs.md)|_Property_ > worksheetId|Represents the id of the worksheet which contains the table.|beta|
|[_WorksheetSelectionChangedEventArgs](reference/excel/_worksheetselectionchangedeventargs.md)|_Property_ > address|Gets the range address that represents the selected area of a specific worksheet.|1.7|
|[_WorksheetSelectionChangedEventArgs](reference/excel/_worksheetselectionchangedeventargs.md)|_Property_ > worksheetId|Gets the range address that represents the selected area of a specific worksheet.|beta|
