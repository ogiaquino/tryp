<%inherit file='base.mako'/>

${form.begin(request.route_url('report-new'))}
    <div class="row-fluid">
        <label for="textinput2">
            Enter new report's file name
        </label>
    </div>
    <div class="row-fluid">
        ${form.text('filename')}
    </div>
    <div class="row-fluid">
        ${form.errorlist('filename')}
    </div>
    <div class="row-fluid">
        ${form.submit('form.submitted', 'save', class_='btn')}
    </div>
${form.end()}
