<?php
// with Composer
require_once 'vendor/autoload.php';

use AlejandroSosa\YiiPowerPoint\PowerPoint;

//setting file
$options = [
    //file
    'fileName' => Yii::t('app', 'demo'),

    //properties of file
    'fileProperties' => [
        'creator' => Yii::t('app', 'MyCompany'),
        'title' => Yii::t('app', 'Financial Statement 2016'),
        'subject' => Yii::t('app', 'General resume'),
        'description' => Yii::t('app', 'Budget report')
    ],

    //design all ppt
    'layout' => [
        'background' => Yii::getPathOfAlias('images') .'/bg.png'
    ],
];

//slides of ppt
$slides = [
    [
        //array of text
        'texts' => [
            [
                'text'=>Yii::t('app', 'Delivery'),
                'options'=>[
                    'height'=>82, 'width'=>816, 'ox'=>72, 'oy'=>170,
                    'bold'=> true, 'size'=>44, 'color'=>'00000000',
                    'align'=>PowerPoint::TEXT_ALIGN_HORIZONTAL_CENTER,
                ],
            ],
            [
                'text'=>Yii::t('app', 'State Budget 2016'),
                'options'=>[
                    'height'=>82, 'width'=>816,
                    'ox'=>72, 'oy'=>291,
                    'bold'=> false, 'size'=>40, 'color'=>'00000000',
                    'align'=>PowerPoint::TEXT_ALIGN_HORIZONTAL_CENTER,
                ],
            ],
        ]
    ]
];

//generate ppt
Yii::app()->ppt->generate($options, $slides);