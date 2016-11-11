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
                    'height'=>38, 'width'=>360, 'ox'=>345, 'oy'=>34,
                    'bold'=> true, 'size'=>18, 'color'=>PowerPoint::COLOR_CYAN,
                    'align'=>PowerPoint::TEXT_ALIGN_VERTICAL_CENTER,
                ],
            ],
            [
                'text'=>Yii::t('app', 'Abril/2016'),
                'options'=>[
                    'height'=>38, 'width'=>132, 'ox'=>816, 'oy'=>52,
                    'bold'=> true, 'size'=>18, 'color'=>PowerPoint::COLOR_CYAN,
                    'align'=>PowerPoint::TEXT_ALIGN_VERTICAL_CENTER,
                ],
            ],
            [
                'text'=>Yii::t('app', 'State Budget 2016'),
                'options'=>[
                    'height'=>38, 'width'=>403, 'ox'=>326, 'oy'=>73,
                    'bold'=> true, 'size'=>18, 'color'=>PowerPoint::COLOR_CYAN,
                    'align'=>PowerPoint::TEXT_ALIGN_VERTICAL_CENTER,
                ],
            ],
        ],
        'images' => [
            [
                'options'=>[
                    'path'=> '', 'name'=> '', 'decription'=> '',
                    'height'=>234, 'width'=>450, 'ox'=>18, 'oy'=>117
                ]
            ],
            [
                'options'=>[
                    'path'=> '', 'name'=> '', 'decription'=> '',
                    'height'=>234, 'width'=>450, 'ox'=>487, 'oy'=>117
                ]
            ]
        ],
    ]
];

//generate ppt
Yii::app()->ppt->generate($options, $slides);