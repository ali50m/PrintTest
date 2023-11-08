# PrintTest

### xaml code
```xml
<Grid>
    <Grid.RowDefinitions>
        <RowDefinition Height="Auto" />
        <RowDefinition Height="*" />
    </Grid.RowDefinitions>
    <Button Grid.Row="0" Click="OnPrintClick">Print</Button>
    <Canvas Name="Canvas" Grid.Row="1">
        <TextBlock
            Canvas.Left="65"
            Canvas.Top="51"
            Text="Hello There" />
    </Canvas>
</Grid>
```

### invoke the PrintDialog
```csharp
private void OnPrintClick(object sender, System.Windows.RoutedEventArgs e)
{
    var printDialog = new PrintDialog();

    if (printDialog.ShowDialog() == true)
    {
        printDialog.PrintVisual(Canvas, "A Simple Drawing");
    }
}
```

### screenshots

![image](https://github.com/ali50m/PrintTest/assets/9393831/efeaf55c-5b1d-41d1-9844-2ea5b9c72269)

![image](https://github.com/ali50m/PrintTest/assets/9393831/e24b3d49-0218-4a3b-aff0-4e588f194b29)

![image](https://github.com/ali50m/PrintTest/assets/9393831/48dd1f87-ffa6-4dcf-b26f-5c7201c8534c)

![image](https://github.com/ali50m/PrintTest/assets/9393831/7821c6de-a3b6-4bf7-a5aa-d2f36013c819)
