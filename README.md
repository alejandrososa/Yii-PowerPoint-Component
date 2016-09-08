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
                ...
                'application.vendor.*',
                ...
            )
            ...
        ),

USAGE
-----

The first step in our controller is import Power Point class

        // in your Controller
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
                            'height'=>143, 'width'=>793,
                            'ox'=>120, 'oy'=>180,
                            'bold'=> true, 'size'=>44, 'color'=>'00000000',
                            'align'=>PowerPoint::TEXT_ALIGN_HORIZONTAL_CENTER,
                        ],
                        [
                            'text'=>Yii::t('app', 'State Budget 2016'),
                            'height'=>143, 'width'=>793,
                            'ox'=>170, 'oy'=>240,
                            'bold'=> false, 'size'=>40, 'color'=>'00000000',
                            'align'=>PowerPoint::TEXT_ALIGN_HORIZONTAL_CENTER,
                        ],
                    ]
                ]
            ];
            
            $objPPT = new PowerPoint();
            $objPPT->options = $options;
            $objPPT->slides = $slides;
            $objPPT->exportPPT();
        }
       
       

WHAT'S NEXT
-----------

If you notice a bug, think about some improvement, then please write me