function myFunction() {

  // 現在利用中のシート
  var sheet = SpreadsheetApp.getActiveSheet();

  // getRange(row, column)

  var range;
  var val;
  var column;
  var row;
  var tpoint = 0;

  // 最後の問題の次のセルにセル数をセット
  for( column = 1; column < 1000; column++ ) {
    range = sheet.getRange(1, column);
    val = range.getValue().toString();

    switch (column) {
      case 1:
      case 2:
        break;
      case 17:
      case 18:
      case 19:
      case 26:
      case 42:
      case 43:
      case 44:
        tpoint += 5;
        break;
      case 22:
      case 23:
      case 27:
        tpoint += 2;
        break;
      case 33:
      case 36:
      case 37:
      case 38:
      case 39:
      case 40:
      case 41:
        tpoint += 3;
        break;
      default:
        tpoint++;
    }    

    if ( val == "" ) {
      Logger.log( tpoint );
      range.setValue( tpoint );
      break;
    }

  }

  // 学生数をカウント
  for( row = 1; row < 1000; row++ ) {
    range = sheet.getRange(row, 1);
    val = range.getValue().toString();

    if ( val == "" ) {
      Logger.log( row );
      break;
    }

  }

  // 正解色
  var targetColor = sheet.getRange(1, 1).getBackground();
  var ok;

  var i;
  var point = 0;

  // 正解なら 1 ポイントアップ。( 問題番号によって、加算する場合もありえます )
  for ( j = 3; j < row; j++ ) {
    for ( i = 3; i < column; i++ ) {
      ok = sheet.getRange(j, i).getBackground();
      if ( ok == targetColor ) {
        switch (i) {
          case 17:
          case 18:
          case 19:
          case 26:
          case 42:
          case 43:
          case 44:
            point += 5;
            break;
          case 22:
          case 23:
          case 27:
            point += 2;
            break;
          case 33:
          case 36:
          case 37:
          case 38:
          case 39:
          case 40:
          case 41:
            point += 3;
            break;
          default:
            point++;
        }
      }
    }

    // 正解点数
    range = sheet.getRange(j, column);
    range.setValue( point );

    // 100点換算( 切り上げ )
    range = sheet.getRange(j, column + 1);
    range.setValue( Math.ceil( point * 100 / (tpoint) ) );

    // 点数の後ろに名前の転送
    range = sheet.getRange(j, 2);
    val = range.getValue().toString();
    range = sheet.getRange(j, column + 2);
    range.setValue( val );

    // 点数の初期化
    point = 0;
  }

}


