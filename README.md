Yii PowerPoint Component
========================

Yii PowerPoint - Create and write presentations, with slides; slides are made up of text, images in PHP

[![Build Status](https://travis-ci.org/alejandrososa/Yii-PowerPoint-Component.svg?branch=master)](https://travis-ci.org/alejandrososa/Yii-PowerPoint-Component)

QUICK START
-----------

You can see all examples and documentation:

      yiiAppPath/protected/vendor/alejandrososa/yii-power-point/samples
      yiiAppPath/protected/vendor/alejandrososa/yii-power-point/README.md

INSTALLATION
------------

On command line, type in the following commands:

        $ cd yiiAppPath/protected              
        $ composer init && composer install                    (optional if not have composer.json)
        $ composer require alejandrososa/yii-power-point dev-master

Assuming that the vendor is installed in yiiAppPath/index.php

        // Include the vendor's autoloader:
        $vendor = dirname(__FILE__) . '/protected/vendor/autoload.php';
        require_once($vendor);

Now we import the vendor installed in the library main.php
Edit the protected/config/main.php adding the following:

        //(optional path alias for images)
        Yii::setPathOfAlias('images', dirname(dirname(dirname(__FILE__))).'/images');
        
        return array(
            // autoloading model and component classes
            'import' => array(
                'application.vendor.*',
                ...
            )
            // application components
            'components'=>array(
                'ppt' => array(
                    'class' => 'AlejandroSosa\\YiiPowerPoint\\PowerPoint'
                )
                ...
            )
        ),

USAGE
-----

The first step in our controller is import Power Point class

        // in your Controller
        (use optional, only to call constants defined in PowerPoint)
        use AlejandroSosa\YiiPowerPoint\PowerPoint;
        
        class SiteController extends Controller {
            ...
        }
   
Now in your action       
        
        /**
        * Create PPT
        */
        public function actionCreatePPT()
        {
            $options = [
                'fileName' => 'myFirstPPT',
                'fileProperties' => [
                    'creator' => Yii::t('app', 'MyCompany'),
                    'title' => Yii::t('app', 'State Budget 2016'),
                    'subject' => Yii::t('app', 'Resumen General'),
                    'description' => Yii::t('app', 'Any description')
                ],
                'layout' => [
                    'background' => Yii::getPathOfAlias('images') .'/ppt/bg.png'
                ]
            ];
    
            $slides = [
                [
                    'texts' => [
                        [
                            'text'=>Yii::t('app', 'Delivery'),
                            'options'=>[
                                'height'=>143, 'width'=>793, 'ox'=>120, 'oy'=>180,
                                'bold'=> true, 'size'=>44, 'color'=>PowerPoint::COLOR_CYAN,
                                'align'=>PowerPoint::TEXT_ALIGN_HORIZONTAL_CENTER,
                            ],
                        ],
                        [
                            'text'=>Yii::t('app', 'State Budget 2016'),
                            'options'=>[
                                'height'=>143, 'width'=>793, 'ox'=>170, 'oy'=>240,
                                'bold'=> false, 'size'=>40, 'color'=>PowerPoint::COLOR_BLUE,
                                'align'=>PowerPoint::TEXT_ALIGN_HORIZONTAL_CENTER,
                            ],
                        ] 
                    ]
                ]
            ];
            
            Yii::app()->ppt->generate($options, $slides);
        }
       
SETTINGS OF CHARTS    
------------------       

**General Purpose in options**

| Property  | Definition | Type |
|---|---|---|
|height   |height of the graph in pixels| integer |
|width   |width of the graph in pixels| integer |
|ox|X coordinate, horizontal position on the slide| integer |
|oy|Y coordinate, vertical position on the slide| integer |

**IMAGES**

| Property  | Definition | Type|
|---|---|---|
|**options**| Includes the following properties| array |
|path   |Absolute path of the image| string |
|name   |Name of the image (optional)| string |
|decription|Description of the image (optional)| string |

**TEXTS**

| Property  | Definition | Type|
|---|---|---|
|**text**   |Text you want to insert| string |
|**options**| Includes the following properties| array |
|align|Text alignment| string, you can access the constants from PowerPoint::TEXT_ALIGN_HORIZONTAL_CENTER |
|bold|Bold text| boolean, default is false |
|color|Text color| string, default is COLOR_PRIMARY_TEXT. You can access the constants from PowerPoint::COLOR_BLUE |
|size|Text size, default is 14 | integer |

**TABLES**

| Property  | Definition | Type|
|---|---|---|
|height   |Height of the table in pixels| integer |
|width   |Width of the table in pixels| integer |
|ox|X coordinate, horizontal position on the slide| integer |
|oy|Y coordinate, vertical position on the slide| integer |
|**header**   |Row that represents the header of the table, array with the columns it contains| array |
|columns   |Array of columns, contains text and column style| array |
|text   |Text you want to insert into column| string |
|**style**| Includes the following properties into column,| array |
|background|Background color of column| string, default is COLOR_WHITE. You can access the constants from PowerPoint::COLOR_WHITE |
|height   |Height of column in pixels| integer, default is 20 |
|width   |Width of colum in pixels| integer, default is 100 |
|align|Text alignment of column| string default is center, you can access the constants from PowerPoint::TEXT_ALIGN_HORIZONTAL_CENTER |
|bold|Bold text of column| boolean, default is false |
|color|Text color of column| string, default is COLOR_PRIMARY_TEXT. You can access the constants from PowerPoint::COLOR_BLUE |
|size|Text size of column, default is 14 | integer |
|**rows**   |Row that represents the header of the table, array with the columns it contains| array |
|columns   |Array of columns, contains text and column style or only text| array or string|
|**style**| Like the header, you can use the style properties in the rows at the general level for all rows | array |

**Nota**: Styles can be applied at the general level for all rows or at a specific level for a column, you can see sample_tables


**CHARTS**

| Property  | Definition | Type|
|---|---|---|
|**title**   |Title of chart| string |
|**series**| Array of data series composed of key and value array('2015'=>150) | array(string=>integer) |
|**options**| Includes the following properties| array |
|type|Type of chart to create | string, you can access the constants from PowerPoint::TEXT_ALIGN_HORIZONTAL_CENTER |
|seriesName|Names of the data series, the name quantity must be equal to the number of data series| array |
|showSeriesName|Visibility of data series names, default is true | array |
|showPercente|Displays the percentages of the data series, default is true | boolean |
|showValue|Displays the values of the data series, default is true | boolean |
|showSymbol|Show symbol in data series, default is true. Only in Scatter | boolean |
|typeSymbol|Show symbol in data series, default is true. Only in Scatter | boolean |
|bold|Bold text| boolean, default is false |
|color|Text color| string, default is COLOR_PRIMARY_TEXT. You can access the constants from PowerPoint::COLOR_BLUE |
|size|Text size, default is 14 | integer |

Available the following type charts: CHART_TYPE_BAR, CHART_TYPE_BAR_HORIZONTAL, CHART_TYPE_BAR_STACKED, CHART_TYPE_BAR_PERCENT_STACKED, CHART_TYPE_BAR_3D, CHART_TYPE_PIE, CHART_TYPE_PIE_3D, CHART_TYPE_SCATTER 
Available the following symbols in scatter: SYMBOL_CIRCLE, SYMBOL_DASH, SYMBOL_DIAMOND, SYMBOL_DOT, SYMBOL_NONE, SYMBOL_PLUS, SYMBOL_SQUARE, SYMBOL_STAR, SYMBOL_TRIANGLE, SYMBOL_X        


CHART_TYPE_BAR, CHART_TYPE_BAR_HORIZONTAL, CHART_TYPE_BAR_STACKED, CHART_TYPE_BAR_PERCENT_STACKED, CHART_TYPE_BAR_3D, CHART_TYPE_PIE, CHART_TYPE_PIE_3D, CHART_TYPE_SCATTER 


'type'=>PowerPoint::CHART_TYPE_BAR, 'seriesName'=> ['2014', '2015', '2016'], 'showSeriesName'=> false,
                    'titleItalic' => true, 'legendItalic' => true,
                    'height'=>234, 'width'=>450, 'ox'=>18, 'oy'=>117
                    
                    
OTHER EXAMPLES    
--------------

The other examples to create tables and graphs can be found in vendor/alejandrososa/yii-power-point/samples

WHAT'S NEXT
-----------

If you notice a bug, think about some improvement, then please write me