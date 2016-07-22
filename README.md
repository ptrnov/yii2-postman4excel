Postman Excel 
==============
Excel on view or console cronjobs

Installation
------------

The preferred way to install this extension is through [composer](http://getcomposer.org/download/).

Either run

```
php composer.phar require --prefer-dist ptrnov/yii2-postman4excel "dev-master"
```

or add

```
"ptrnov/yii2-postman4excel": "dev-master"
```

to the require section of your `composer.json` file.


Usage
-----

Once the extension is installed, simply use it in your code by  :

```php
<?= \ptrnov\excel\AutoloadExample::widget(); ?>```