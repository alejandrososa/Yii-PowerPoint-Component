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
                'height'=>38, 'width'=>360, 'ox'=>345, 'oy'=>34,
                'bold'=> true, 'size'=>18, 'color'=>'FF2196F3',
                'align'=>PowerPoint::TEXT_ALIGN_VERTICAL_CENTER,
            ],
            [
                'text'=>Yii::t('app', 'Abril/2016'),
                'height'=>38, 'width'=>132, 'ox'=>816, 'oy'=>52,
                'bold'=> true, 'size'=>18, 'color'=>'FF2196F3',
                'align'=>PowerPoint::TEXT_ALIGN_VERTICAL_CENTER,
            ],
            [
                'text'=>Yii::t('app', 'State Budget 2016'),
                'height'=>38, 'width'=>403, 'ox'=>326, 'oy'=>73,
                'bold'=> true, 'size'=>18, 'color'=>'FF2196F3',
                'align'=>PowerPoint::TEXT_ALIGN_VERTICAL_CENTER,
            ],
        ],
        'tables' => [
            [
                'height'=>400, 'width'=>900, 'ox'=>26, 'oy'=>163,
                'header' => [
                    'columns'=> [
                        ['text'=>'Main Title', 'style'=>['width'=>268]],
                        ['text'=>'Title', 'style'=>['width'=>88]],
                        ['text'=>'Title', 'style'=>['width'=>88]],
                        ['text'=>'Title', 'style'=>['width'=>88]],
                        ['text'=>'Title', 'style'=>['width'=>88]],
                        ['text'=>'Title', 'style'=>['width'=>88]],
                        ['text'=>'Title', 'style'=>['width'=>88]],
                        ['text'=>'Title', 'style'=>['width'=>88]],
                    ],
                    'style'=>[
                        'background'=>'FFC6D9F1', 'bold'=> true, 'size'=>9, 'color'=>'FF000000',
                        'align'=>PowerPoint::TEXT_ALIGN_HORIZONTAL_CENTER, 'height'=>35
                    ],
                ],
                'rows'=>[
                    [
                        'columns'=>[
                            //with style
                            ['text'=>'Title of this row', 'style'=>['align'=>PowerPoint::TEXT_ALIGN_HORIZONTAL_LEFT]],
                            //without style
                            '2.140',
                            '26,07%',
                            '9.00',
                            '2.121',
                            '1.554',
                            '7.000',
                            '73,27%'
                        ],
                        'style'=>[
                            'background'=>'FFD7E4BD', 'bold'=> true, 'size'=>10, 'color'=>'FF000000',
                            'align'=>PowerPoint::TEXT_ALIGN_HORIZONTAL_RIGHT, 'height'=>17
                        ]
                    ],
                    [
                        'columns'=>[
                            //with style
                            ['text'=>'Second title of this row', 'style'=>['align'=>PowerPoint::TEXT_ALIGN_HORIZONTAL_LEFT]],
                            //without style
                            '2.140',
                            '26,07%',
                            '9.00',
                            '2.121',
                            '1.554',
                            '7.000',
                            '73,27%'
                        ],
                        'style'=>[
                            'background'=>'FFE9EDF4', 'bold'=> false, 'size'=>10, 'color'=>'FF000000',
                            'align'=>PowerPoint::TEXT_ALIGN_HORIZONTAL_RIGHT, 'height'=>17
                        ]
                    ],
                ],
            ],
        ],
    ]
];

//generate ppt
Yii::app()->ppt->generate($options, $slides);