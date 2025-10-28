odoo.define('supplier_payment_receipts.form_widgets', function(require) {
    "use strict";

    var core = require('web.core');
    var utils = require('web.utils');
    var fieldRegistry = require('web.field_registry');
    var ListRenderer = require('web.ListRenderer');
    var rpc = require('web.rpc');
    var FieldOne2Many = require('web.relational_fields').FieldOne2Many;
    var _t = core._t;
    var FormController = require('web.FormController');
    var core = require('web.core');
    var errorShown = false;

    ListRenderer.include({
        _updateSelection: function () {
            const previousSelection = JSON.stringify(this.selection);
            this.selection = [];
            var self = this;
            var $inputs = this._getSelectableRecordCheckboxes();
            var allChecked = $inputs.length > 0;
            
            $inputs.each(function (index, input) {
                if (input.checked) {
                    self.selection.push($(input).closest('tr').data('id'));
                } else {
                    allChecked = false;
                }
            });
            
            this.$('thead .o_list_record_selector input').prop('checked', allChecked);
            
            if (JSON.stringify(this.selection) !== previousSelection) {
                this.trigger_up('selection_changed', { allChecked, selection: this.selection });
            }
            
            // Actualizar contador y controles usando el conteo correcto
            this._updateSelectionControls(this.selection.length, $inputs.length);
            
            this._updateFooter();
        },

        _updateSelectionControls: function(selectedCount, totalCount) {
            // Usar el totalCount que viene de los checkboxes reales
            var realTotalCount = totalCount;
            
            // Actualizar contador de líneas
            var counterText = selectedCount + ' de ' + realTotalCount + ' seleccionadas';
            $('.selected_count_display').text(counterText);
            
            // Actualizar barra de progreso
            var percentage = realTotalCount > 0 ? (selectedCount / realTotalCount) * 100 : 0;
            $('.selection-progress').css('width', percentage + '%').attr('aria-valuenow', percentage);
            
            // Mostrar/ocultar TODO EL WIDGET según selección
            if (selectedCount > 0) {
                $('.payment-widget-card').show();  // Mostrar todo el widget
                $('.button_select_order_lines').show();
                
                // Cambiar color de la barra según cantidad seleccionada
                if (percentage === 100) {
                    $('.selection-progress').removeClass('bg-warning bg-info').addClass('bg-success');
                } else if (percentage > 50) {
                    $('.selection-progress').removeClass('bg-success bg-info').addClass('bg-warning');
                } else {
                    $('.selection-progress').removeClass('bg-success bg-warning').addClass('bg-info');
                }
                
                // Actualizar texto de ayuda
                $('.quick-actions span').html('<i class="fa fa-check-circle mr-1 text-success"/> ' + selectedCount + ' líneas listas para procesar');
                
            } else {
                $('.payment-widget-card').hide();  // Ocultar todo el widget
                $('.button_select_order_lines').hide();
                $('.selection-progress').removeClass('bg-success bg-warning bg-info').addClass('bg-secondary');
                $('.quick-actions span').html('<i class="fa fa-info-circle mr-1"/> Selecciona las líneas para continuar');
            }
        }
    });

    var TotalDisplayWidget = FieldOne2Many.extend({
        template: 'PaymentReceiptWidget',
        events: {
            "keyup .oe_search_value": "_onKeyUp",
            "click .button_select_order_lines": "selected_lines",
            "click": "updateAmount",
        },
        
        init: function() {
            var self = this;
            this._super.apply(this, arguments);
        },

        start: function () {
            var self = this;
            return this._super.apply(this, arguments).then(function () {
                // Inicializar contador
                self._initializeCounters();
                // Primer cálculo
                setTimeout(function() {
                    self.updateAmount();
                    self._updateTotalCounter();
                }, 100);
            });
        },

        _initializeCounters: function() {
            $('.selected_count_display').text('0 de 0 seleccionadas');
            $('.total_amount_display').text('$ 0,00');
            $('.selection-progress').css('width', '0%');
            $('.payment-widget-card').hide(); // Ocultar widget al inicializar
        },

        _updateTotalCounter: function() {
            var self = this;
            // Contar solo las filas que tienen checkbox selector (filas de datos reales)
            var totalRows = this.$el.find('td.o_list_record_selector').length;
            var selectedRows = this.find_selected_lines().length;
            
            // Actualizar contador
            var counterText = selectedRows + ' de ' + totalRows + ' seleccionadas';
            $('.selected_count_display').text(counterText);
            
            // Actualizar barra de progreso
            var percentage = totalRows > 0 ? (selectedRows / totalRows) * 100 : 0;
            $('.selection-progress').css('width', percentage + '%').attr('aria-valuenow', percentage);
            
            // Mostrar/ocultar TODO EL WIDGET según selección
            if (selectedRows > 0) {
                $('.payment-widget-card').show();  // Mostrar todo el widget
            } else {
                $('.payment-widget-card').hide();  // Ocultar todo el widget
            }
            
            // Cambiar estilos según progreso
            if (percentage === 100) {
                $('.selection-progress').removeClass('bg-warning bg-info bg-secondary').addClass('bg-success');
            } else if (percentage > 50) {
                $('.selection-progress').removeClass('bg-success bg-info bg-secondary').addClass('bg-warning');
            } else if (percentage > 0) {
                $('.selection-progress').removeClass('bg-success bg-warning bg-secondary').addClass('bg-info');
            } else {
                $('.selection-progress').removeClass('bg-success bg-warning bg-info').addClass('bg-secondary');
            }
        },

        _onKeyUp: function (event) {
            var value = $(event.currentTarget).val().toLowerCase();
            this.$el.find('table').addClass('oe_one2many');
            $(".oe_one2many tr:not(:lt(1))").filter(function () {
                $(this).toggle($(this).text().toLowerCase().indexOf(value) > -1)
            });
            
            // Actualizar contador después del filtro
            setTimeout(() => {
                this._updateTotalCounter();
            }, 100);
        },        
        
        selected_lines: function() {
            var self = this;
            var current_model = this.recordData[this.name].model;
            var selected_lines = self.find_unselected_lines();
            var selected_lines_remove = self.find_selected_lines();
        
            if (selected_lines_remove.length === 0) {
                return this.displayNotification({
                    message: _t('Seleccione al menos un registro.'),
                    type: 'danger'
                });
            }

            // Mostrar indicador de procesamiento
            $('.processing-indicator').show();
            $('.button_select_order_lines').prop('disabled', true);
        
            var w_response = confirm("¿Quieres confirmar los " + selected_lines_remove.length + " registros seleccionados?");
            if (w_response) {
                var target_model;
                var method;
        
                if (current_model === 'supplier.payment.receipt.line') {
                    target_model = 'supplier.payment.receipt';
                    method = 'remove_lines';
                } else {
                    console.log("Model not matched, exiting function.");
                    $('.processing-indicator').hide();
                    $('.button_select_order_lines').prop('disabled', false);
                    return;
                }
        
                rpc.query({
                    'model': target_model,
                    'method': method,
                    'args': [selected_lines, selected_lines_remove]
                }).then(function(result) {
                    $('.processing-indicator').hide();
                    $('.button_select_order_lines').prop('disabled', false);
                    self.trigger_up('reload');
                }).catch(function(error) {
                    $('.processing-indicator').hide();
                    $('.button_select_order_lines').prop('disabled', false);
                    self.displayNotification({
                        message: _t('Error al procesar: ') + error.message,
                        type: 'danger'
                    });
                });
            } else {
                $('.processing-indicator').hide();
                $('.button_select_order_lines').prop('disabled', false);
            }
        },

        _getRenderer: function() {
            if (this.view.arch.tag === 'tree') {
                return ListRenderer.extend({
                    init: function(parent, state, params) {
                        this._super.apply(this, arguments);
                        this.hasSelectors = true;
                    },
                });
            }
            return this._super.apply(this, arguments);
        },

        find_selected_lines: function() {
            var self = this;
            var selected_list = [];
            this._getRenderer();
            this.$el.find('td.o_list_record_selector').each(function() {
                var record_id = parseInt(self._getResId($(this).closest('tr').data('id')));
                if ($(this).find('input:checked').length > 0) {
                    selected_list.push(record_id);
                }
            });
            return selected_list;
        },

        find_unselected_lines: function() {
            var self = this;
            var unselected_list = [];
            this.$el.find('td.o_list_record_selector').each(function() {
                var record_id = parseInt(self._getResId($(this).closest('tr').data('id')));
                if (!$(this).find('input:checked').length) {
                    unselected_list.push(record_id);
                }
            });
            return unselected_list;
        },

        _getResId: function(recordId) {
            var record;
            utils.traverse_records(this.recordData[this.name], function(r) {
                if (r.id === recordId) {
                    record = r;
                }
            });
            if (!record || !record.res_id) {
                if (!errorShown) {
                    this.displayNotification({
                        message: _t('Por favor edite cambie la paginación y guarde el registro!'),
                        type: 'danger'
                    });
                    errorShown = true;
                }
                return false;
            }
            return record.res_id;
        },

        updateAmount: function() {
            var self = this;
            var current_model = this.recordData[this.name].model;
            var selected_lines = self.find_selected_lines();
        
            if (selected_lines.length === 0) {
                $('.total_amount_display').text('$ 0,00');
                this._updateTotalCounter();
                return;
            }
        
            var target_model;
            var method;
        
            if (current_model === 'supplier.payment.receipt.line') {
                target_model = 'supplier.payment.receipt';
                method = 'update_amount';
            } else {
                console.log("Model not matched, exiting function.");
                return;
            }
        
            rpc.query({
                'model': target_model,
                'method': method,
                'args': [selected_lines],
            }).then(function(totalAmount) {
                if (totalAmount === undefined) {
                    totalAmount = 0;
                }
                var formattedAmount = self._formatCurrency(totalAmount);
                $('.total_amount_display').text(formattedAmount);
                
                // Actualizar contador después de calcular
                self._updateTotalCounter();
            }).catch(function(error) {
                console.error('Error calculating amount:', error);
                $('.total_amount_display').text('$ 0,00');
                self._updateTotalCounter();
            });
        },

        _formatCurrency: function(amount) {
            return new Intl.NumberFormat('es-CO', {
                style: 'currency',
                currency: 'COP',
                minimumFractionDigits: 2
            }).format(amount);
        },
    });
    
    fieldRegistry.add('payment_supplier_widget', TotalDisplayWidget);
});